using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Timers;
using System.Globalization;

namespace PowerLib
{
    public class PowerSupply
    {
        public const decimal MAX_VOLTS = 30;
        public const decimal MAX_AMPS = 3;
        public const int MEMORY_RANGE_LOW = 1;     // Lowest memory location to save settings to on the supply
        public const int MEMORY_RANGE_HIGH = 4;    // Highest memory location to save settings to on the supply
        public const int TRACK_MODE_MIN = 0;       // Minimum tracking mode
        public const int TRACK_MODE_MAX = 2;       // Maximum tracking mode
        public const string MODEL = "GW INSTEK,GPD-3303S";
        const int TIME_WAIT_RECEIVE = 700;  // Maximum time to wait for serial receive
        const int DATA_OFFSET_END = 2;      // Amount of characters to remove to get decimal-only values
        string[] ports = SerialPort.GetPortNames();
        string sLatestData;
        bool bDataFlag = false;     // More data received?
        bool bWatchdog = false;     // Taking too much time?
        bool bBeepStatus = true;    // Current status of beep? (assume on)
        SerialPort powerPort = new SerialPort();
        Timer watchdog = new Timer(TIME_WAIT_RECEIVE);

        /// <summary>
        /// Constructor for the power supply, initializes connection
        /// </summary>
        public PowerSupply()
        {
            bool bFoundPort = false;

            // Make sure port is closed
            powerPort.Close();

            // Intialize serial and timer event handlers
            powerPort.DataReceived += new SerialDataReceivedEventHandler( dataReceiveHandler );
            watchdog.Elapsed += new ElapsedEventHandler( watchdogHandler );

            // Initialize serial settings
            powerPort.ReadTimeout = 2000000;
            powerPort.WriteTimeout = 2000000;
            powerPort.BaudRate = 9600;
            powerPort.Parity = Parity.None;
            powerPort.DataBits = 8;
            powerPort.StopBits = StopBits.One;
            powerPort.Handshake = Handshake.None;
            powerPort.RtsEnable = false;

            bFoundPort = false;
            foreach ( string port in ports ) {
                // Intialize serial settings
                powerPort.PortName = port;
                try {
                    powerPort.Open();
                }
                catch {
                    // Could not grab port
                    continue;
                }

                // Check if this port is connected to supply
                if ( SupplyCheck() ) {
                    bFoundPort = true;
                    break;
                }
                powerPort.Close();
            }
            if ( !bFoundPort ) {
                throw new NoSupplyException( "Failed to connect to power supply" );
            } else {
                // Ensure supply is turned off
                TurnOffPower();
            }
        }

        /// <summary>
        /// Checks if there is a connection with the power supply
        /// </summary>
        /// <returns>true on success, false on failure</returns>
        public bool SupplyCheck()
        {
            string sModel;

            sModel = getResponse( "*IDN?" );
            if ( sModel != null ) {
                if ( sModel.StartsWith( MODEL ) ) {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Set the maximum current on specific channel
        /// </summary>
        /// <param name="iChnl">The channel selection to first initialize on the supply (1 or 2)</param>
        /// <param name="dCrnt">The current level in amperes to set on the power supply (0.001A to 3A)</param>
        /// <returns>0 on success, -1 on failure</returns>
        public int SetCurrent(int iChnl, decimal dCrnt)
        {
            // Check channel and current level
            if( iChnl < 1 || iChnl > 2 ) {
                Console.WriteLine( "Invalid channel, must be 1 or 2" );
                return -1;
            } else if ( dCrnt <= 0 || dCrnt > MAX_AMPS ) {
                Console.WriteLine( "Maximum current cannot be less than/equal to 0 or greater than 3A" );
                return -1;
            }

            // Compose string and send command, check if succeeded
            string sCmd = "ISET" + iChnl + ":" + dCrnt;
            powerPort.WriteLine( sCmd );

            if ( Decimal.Compare( dCrnt, GetCurrent( iChnl ) ) == 0 ) {
                return 0;
            } else {
                return -1;
            }
        }

        /// <summary>
        /// Set the voltage on specific channel
        /// </summary>
        /// <param name="iChnl">The channel selection to first initialize on the supply (1 or 2)</param>
        /// <param name="dVolt">The voltage level in volts to set on the specified channel (0.001V to 30V)</param>
        /// <returns>0 on success, -1 on failure</returns>
        public int SetVoltage(int iChnl, decimal dVolt)
        {
            // Check channel and current level
            if (iChnl < 1 || iChnl > 2) {
                Console.WriteLine( "Invalid channel, must be 1 or 2" );
                return -1;
            } else if ( dVolt <= 0 || dVolt > MAX_VOLTS ) {
                Console.WriteLine( "Voltage cannot be less than/equal to 0 or greater than 30V" );
                return -1;
            }

            // Compose string and send command, check if succeeded
            string sCmd = "VSET" + iChnl + ":" + dVolt;
            powerPort.WriteLine( sCmd );

            if ( Decimal.Compare( dVolt, GetVoltage( iChnl ) ) == 0 ) {
                return 0;
            } else {
                return -1;
            }
        }

        /// <summary>
        /// Finds the set maximum current reading
        /// </summary>
        /// <param name="iChnl">The channel selection to first initialize on the supply (1 or 2)</param>
        /// <returns>set value of current in amperes on success, max current + 1 on failure</returns>
        public decimal GetCurrent(int iChnl)
        {
            decimal dReturn = MAX_AMPS + 1;
            string sCmd = "ISET" + iChnl + "?";
            string sCurrent;

            sCurrent = getResponse( sCmd );
            if ( sCurrent != null ) {
                sCurrent = sCurrent.Substring( 0, sCurrent.Length-1 - DATA_OFFSET_END );
                if ( !Decimal.TryParse( sCurrent, out dReturn ) ) {
                    throw new ParseValueException( "Failed to parse value for set current" );
                }
            }

            return dReturn;
        }

        /// <summary>
        /// Finds the set voltage reading
        /// </summary>
        /// <param name="iChnl">The channel selection to first initialize on the supply (1 or 2)</param>
        /// <returns>set value of voltage in volts on success, max voltage + 1 on failure</returns>
        public decimal GetVoltage(int iChnl)
        {
            decimal dReturn = MAX_VOLTS + 1;
            string sCmd = "VSET" + iChnl + "?";
            string sVoltage;

            sVoltage = getResponse( sCmd );
            if ( sVoltage != null ) {
                sVoltage = sVoltage.Substring( 0, sVoltage.Length-1 - DATA_OFFSET_END );
                if ( !Decimal.TryParse( sVoltage, out dReturn ) ) {
                    throw new ParseValueException( "Failed to parse value for set voltage" );
                }
            }

            return dReturn;
        }

        /// <summary>
        /// Find the actual current reading
        /// </summary>
        /// <param name="iChnl">The channel selection to first initialize on the supply (1 or 2)</param>
        /// <returns>current being supplied in amperes, max current + 1 on failure</returns>
        public decimal GetActualCurrent(int iChnl)
        {
            decimal dReturn = MAX_AMPS + 1;
            string sCmd = "IOUT" + iChnl + "?";
            string sCurrent;

            sCurrent = getResponse( sCmd );
            if ( sCurrent != null ) {
                sCurrent = sCurrent.Substring( 0, sCurrent.Length-1 - DATA_OFFSET_END );
                if ( !Decimal.TryParse( sCurrent, out dReturn ) ) {
                    throw new ParseValueException( "Failed to parse value for actual current" );
                }
            }
            return dReturn;
        }

        /// <summary>
        /// Find the actual voltage reading
        /// </summary>
        /// <param name="iChnl">The channel selection to first initialize on the supply (1 or 2)</param>
        /// <returns>voltage being supplied in volts, max voltage + 1 on failure</returns>
        public decimal GetActualVoltage(int iChnl)
        {
            decimal dReturn = MAX_VOLTS + 1;
            string sCmd = "VOUT" + iChnl + "?";
            string sVoltage;

            sVoltage = getResponse( sCmd );
            if ( sVoltage != null ) {
                sVoltage = sVoltage.Substring( 0, sVoltage.Length-1 - DATA_OFFSET_END );
                if ( !Decimal.TryParse( sVoltage, out dReturn ) ) {
                    throw new ParseValueException( "Failed to parse value for actual voltage" );
                }
            }

            return dReturn;
        }

        /// <summary>
        /// Save current supply settings to supply memory
        /// Can save to locations 1 through 4, returns 0 on success
        /// </summary>
        /// <param name="iLocation">Memory location to save settings to (1 through 4)</param>
        /// <returns>0 on success, -1 on failure</returns>
        public int SavePowerSettings(int iLocation)
        {
            string sCmd;

            if ( iLocation > MEMORY_RANGE_HIGH || iLocation < MEMORY_RANGE_LOW ) {
                Console.WriteLine( "Invalid parameter to SavePowerSettings, value must be {0} to {1}, got {2}",
                                    MEMORY_RANGE_LOW, MEMORY_RANGE_HIGH, iLocation );
                return -1;
            }

            sCmd = "SAV" + iLocation;
            powerPort.WriteLine( sCmd );

            return 0;
        }

        /// <summary>
        /// Load a previously saved power supply setting from supply memory
        /// Can load from locations 1 through 4, returns 0 on success
        /// </summary>
        /// <param name="iLocation">Memory location to save settings to (1 through 4)</param>
        /// <returns>0 on success, -1 on failure</returns>
        public int LoadPowerSettings(int iLocation)
        {
            string sCmd;

            if ( iLocation > MEMORY_RANGE_HIGH || iLocation < MEMORY_RANGE_LOW ) {
                Console.WriteLine( "Invalid parameter to LoadPowerSettings, value must be {0} to {1}, got {2}",
                                    MEMORY_RANGE_LOW, MEMORY_RANGE_HIGH, iLocation );
                return -1;
            }

            sCmd = "RCL" + iLocation;
            powerPort.WriteLine( sCmd );

            return 0;
        }

        /// <summary>
        /// For input 0, set power supply channels 1 and 2 to be independant
        /// For input 1, set power supply channels 1 and 2 to be tracking (series)
        /// For input 2, set power supply channels 1 and 2 to be tracking (parallel)
        /// </summary>
        /// <param name="iMode">The above specified tracking mode (0 to 2)</param>
        /// <returns> 0 on success, -1 on failure</returns>
        public int SetTrack(int iMode)
        {
            string sCmd;

            if ( iMode < TRACK_MODE_MIN || iMode > TRACK_MODE_MAX ) {
                Console.WriteLine( "Invalide mode for function SetTrack, must be between {0} and {1}, got {2}",
                                    TRACK_MODE_MIN, TRACK_MODE_MAX, iMode );
                return -1;
            }

            sCmd = "TRACK" + iMode;
            powerPort.WriteLine( sCmd );

            return 0;
        }

        /// <summary>
        /// Turn off output power of supply
        /// </summary>
        public void TurnOffPower()
        {
            string sCmd = "OUT0";
            powerPort.WriteLine( sCmd );
        }

        /// <summary>
        /// Turn on output power of the supply
        /// </summary>
        public void TurnOnPower()
        {
            string sCmd = "OUT1";
            powerPort.WriteLine( sCmd );
            // Make a beep as a safety precaution, let people know it is turning on with audible sound
            MakeBeep();
        }

        /// <summary>
        /// Turn off beep
        /// </summary>
        public void TurnOffBeep()
        {
            string sCmd = "BEEP0";

            bBeepStatus = false;
            powerPort.WriteLine( sCmd );
        }

        /// <summary>
        /// Turn on beep
        /// </summary>
        public void TurnOnBeep()
        {
            string sCmd = "BEEP1";

            bBeepStatus = true;
            powerPort.WriteLine( sCmd );
        }

        /// <summary>
        /// Make a beep
        /// </summary>
        public void MakeBeep()
        {
            TurnOffBeep();
            TurnOnBeep();
            // Keep old setting
            if ( !bBeepStatus ) {
                TurnOffBeep();
            }
        }

        /// <summary>
        /// Prints out error messages from the power supply
        /// </summary>
        public void PrintErrors()
        {
            bDataFlag = false;
            string sError;

            sError = getResponse( "ERR?" );
            if ( sError != null ) {
                Console.WriteLine( sError );
            }
        }

        /// <summary>
        /// Function to send a message and get response
        /// Returns null if no response
        /// </summary>
        /// <param name="sCmd">The command to send before getting a response from the supply</param>
        /// <returns>the power supply response on success, null on watchdog timeout</returns>
        private string getResponse(string sCmd)
        {
            string sReturn = null;

            bDataFlag = false;
            bWatchdog = false;
            powerPort.WriteLine( sCmd );

            watchdog.Enabled = true;
            while ( !bDataFlag && !bWatchdog );
            watchdog.Enabled = false;

            if ( bDataFlag ) {
                sReturn = sLatestData;
            } else if ( bWatchdog ) {
                Console.WriteLine( "Power supply -> Watchdog fired" );
            }

            return sReturn;
        }

        /// <summary>
        /// Handler for receiving data from the power supply
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataReceiveHandler(Object sender, SerialDataReceivedEventArgs e)
        {
            int iIndex;

            SerialPort spSender = (SerialPort)sender;
            // Catenate new data and in case of newline earlier in string, cut off old message
            sLatestData = String.Format( "{0}{1}", sLatestData, spSender.ReadExisting() );
            iIndex = sLatestData.IndexOf( '\n' );
            if ( iIndex != sLatestData.Length - 1 ) {
                // Index in middle of string (not end)? Append new data or make new string
                sLatestData = sLatestData.Substring( iIndex + 1 );
                iIndex = sLatestData.IndexOf( '\n' );
            }
            if(sLatestData.Length != 0 && iIndex != -1){
                // Index at end of non-empty string? Data found
                bDataFlag = true;
            }
        }

        /// <summary>
        /// Handler for watchdog
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private void watchdogHandler(Object source, ElapsedEventArgs e)
        {
            bWatchdog = true;
        }

        /// <summary>
        /// Destructor for power supply, turn off output and close port
        /// </summary>
        ~PowerSupply()
        {
            try {
                TurnOffPower();
            }
            catch { 
                // Ignore
            }
            powerPort.Close();
        }
    }

    /// <summary>
    /// Exception for no supply connected
    /// </summary>
    [Serializable]
    public class NoSupplyException : Exception
    {
        public NoSupplyException() { }
        public NoSupplyException(string message) : base( message ) { }
        public NoSupplyException(string message, Exception inner) : base( message, inner ) { }
        protected NoSupplyException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base( info, context ) { }
    }

    [Serializable]
    public class ParseValueException : Exception
    {
        public ParseValueException() { }
        public ParseValueException(string message) : base( message ) { }
        public ParseValueException(string message, Exception inner) : base( message, inner ) { }
        protected ParseValueException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base( info, context ) { }
    }
}
