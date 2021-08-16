using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace Utilities.PowerSupply {
public class PS2000 : IDisposable, IPowerSupply {
    /// NOTA BENE
    /// For EA-PS 2042-20B the max voltage is 40V
    //set verbose to True to see all bytes
    public bool Verbose = true;
    public const int PS_QUERY = 0x40;
    public const int PS_SEND = 0xC0;
    //nominal values, required for all voltage and current calculations
    public float u_nom = 0;
    public float i_nom = 0;
    ///PRIVATE VARS///
    private SerialPort ser;
    ///

    public PS2000(string COMport) {
        ser = new SerialPort(COMport, 115200, Parity.Odd);
        ser.Open();
        u_nom = GetNominalVoltage();
        i_nom = GetNominalCurrent();
        SetRemote(true);
    }

    ~PS2000() {
    }

    public void Dispose() {
        Close();
    }

    public void Close() {
        SetRemote(false);
        ser.Close();
    }

    /// <summary>
    /// Construct telegram
    /// </summary>
    private byte[] Construct(int type, int node, int obj, byte[] data) {
        var telegram = new List<byte>();
        telegram.Add(Convert.ToByte(0x30 + type));// SD (start delimiter)
        telegram.Add(Convert.ToByte(node));// DN (device node)
        telegram.Add(Convert.ToByte(obj));// DN (device node)
        if (data.Length > 0) {
            foreach (char c in data) {
                telegram.Add(Convert.ToByte(c));
            }
            telegram[0] += Convert.ToByte(data.Length - 1);
        }
        var cs = CalculateChecksum(telegram.ToArray());
        telegram.Add(Convert.ToByte((cs >> 8))); //CS0
        telegram.Add(Convert.ToByte((cs & 0xFF))); //CS1 (checksum)
        return telegram.ToArray();
    }

    /// <summary>
    ///  Algorithm to calculate the check sum
    /// </summary>
    private Int16 CalculateChecksum(byte[] data) {
        Int16 cs = 0;
        foreach (byte b in data) {
            cs += b;
        }
        return cs;
    }

    /// <summary>
    /// Compare checksum with header and data in response from device
    /// </summary>
    private bool CheckChecksum(byte[] ans) {
        var ansWithoutChecksum = ans.Slice(0, -2);
        Int16 cs = CalculateChecksum(ansWithoutChecksum);
        if (ans[ans.Length - 2] != Convert.ToByte(cs >> 8) ||
            ans[ans.Length - 1] != Convert.ToByte(cs & 0xFF)) {
            throw new Exception("ERROR: checksum mismatch");
        } else
            return true;
    }

    /// <summary>
    /// Check for errors in response from device
    /// </summary>
    private bool CheckError(byte[] ans) {
        if (ans[2] != 0xFF)
            return true;
        if (ans[3] == 0x00)
            return true;
        else if (ans[3] == 0x03)
            throw new Exception("ERROR: checksum incorrect");
        else if (ans[3] == 0x04)
            throw new Exception("ERROR: start delimiter incorrect");
        else if (ans[3] == 0x05)
            throw new Exception("ERROR: wrong address for output");
        else if (ans[3] == 0x07)
            throw new Exception("ERROR: object not defined");
        else if (ans[3] == 0x08)
            throw new Exception("ERROR: object length incorrect");
        else if (ans[3] == 0x09)
            throw new Exception("ERROR: access denied");
        else if (ans[3] == 0x0F)
            throw new Exception("ERROR: device is locked");
        else if (ans[3] == 0x30)
            throw new Exception("ERROR: upper limit exceeded");
        else if (ans[3] == 0x31)
            throw new Exception("ERROR: lower limit exceeded");
        return false;
    }

    /// <summary>
    /// Send one telegram, receive and check one response
    /// </summary>
    private byte[] Transfer(int type, int node, int obj, byte[] data) {
        byte[] telegram = Construct(type, 0, obj, data);
        if (Verbose) {
            Console.WriteLine("telegram: " + BitConverter.ToString(telegram));
        }
        //send telegram
        ser.Write(telegram, 0, telegram.Length);
        Thread.Sleep(25);//wait 25milisec to ans
        byte[] ans = new byte[ser.BytesToRead];
        ser.Read(ans, 0, ans.Length);
        if (Verbose) {
            Console.WriteLine("answer: " + BitConverter.ToString(ans));
        }
        //check answer
        CheckChecksum(ans);
        CheckError(ans);
        return ans;
    }

    /// <summary>
    /// Get a binary object
    /// </summary>
    private byte[] GetBinary(int obj) {
        byte[] ans = Transfer(PS_QUERY, 0, obj, new byte[] { });
        return ans.Slice(3, -2);
    }

    /// <summary>
    /// Set a binary object
    /// </summary>
    private byte[] SetBinary(int obj, byte mask, byte data) {
        byte[] ans = Transfer(PS_SEND, 0, obj, new byte[] { mask, data });
        return ans.Slice(3, -2);
    }

    /// <summary>
    /// Get a string object
    /// </summary>
    private string GetString(int obj) {
        byte[] ans = Transfer(PS_QUERY, 0, obj, new byte[] { });
        return System.Text.Encoding.UTF8.GetString(ans.Slice(3, -3));
    }

    /// <summary>
    /// Get a float-type object
    /// </summary>
    private float GetFloat(int obj) {
        byte[] ans = Transfer(PS_QUERY, 0, obj, new byte[] { });
        ans = ans.Slice(3, -2);
        Array.Reverse(ans);
        return System.BitConverter.ToSingle(ans, 0);
    }

    /// <summary>
    /// Get an integer object
    /// </summary>
    private int GetInteger(int obj) {
        byte[] ans = Transfer(PS_QUERY, 0, obj, new byte[] { });
        return (ans[3] << 8) + ans[4];
    }

    /// <summary>
    /// Set an integer object
    /// </summary>
    private int SetInteger(int obj, int data) {
        byte[] ans = Transfer(PS_SEND, 0, obj, new byte[] { Convert.ToByte(data >> 8), Convert.ToByte(data & 0xFF) });
        return (ans[3] << 8) + ans[4];
    }

    // object 0
    public string GetTypePS() {
        return GetString(0);
    }

    //object 1
    public string GetSerial() {
        return GetString(1);
    }

    //object 2
    public float GetNominalVoltage() {
        return GetFloat(2);
    }

    //object 3
    public float GetNominalCurrent() {
        return GetFloat(3);
    }

    //object 4
    public float GetNominalPower() {
        return GetFloat(4);
    }

    //object 6
    public string GetArticle() {
        return GetString(6);
    }

    //object 8
    public string GetManufacturer() {
        return GetString(8);
    }

    //object 9
    public string GetVersion() {
        return GetString(9);
    }

    //object 19
    public int GetDeviceClass() {
        return GetInteger(38);
    }

    //object 38
    public float GetOVPThreshold() {
        int v = GetInteger(38);
        return u_nom * v / 25600;
    }

    public int SetOVPThreshold(int u) {
        return SetInteger(38, u);
    }

    //object 39
    public float GetOCPThreshold() {
        int i = GetInteger(39);
        return i_nom * i / 25600;
    }

    public int SetOCPThreshold(int i) {
        return SetInteger(39, i);
    }

    //object 50
    public float GetVoltageSetpoint() {
        int v = GetInteger(50);
        return u_nom * v / 25600;
    }

    public float SetVoltage(double u) {
        return SetInteger(50, Convert.ToInt16((u * 25600.0) / u_nom));
    }

    //object 51
    public float GetCurrentSetpoint() {
        int i = GetInteger(51);
        return i_nom * i / 25600;
    }

    public float SetCurrent(double i) {
        return SetInteger(51, Convert.ToInt16((i * 25600.0) / i_nom));
    }

    //object 54
    public Dictionary<string, bool> GetControl() {
        byte[] ans = GetBinary(54);
        var control = new Dictionary<string, bool>() {
            {"output", false},
            {"remote", false}
        };
        if (ans[1] == 0x01)
            control["output"] = true;
        else
            control["output"] = false;
        if (ans[0] == 0x01)
            control["remote"] = true;
        else
            control["remote"] = false;
        return control;
    }

    private bool SetControl(byte mask, byte data) {
        byte[] ans = SetBinary(54, mask, data);
        //return True if command was acknowledged ("error 0")
        return ans[0] == 0xFF && ans[1] == 0x00;
    }

    public bool GetRemote() {
        return GetControl()["remote"];
    }

    public bool SetRemote(bool remote) {
        if (remote)
            return SetControl(0x10, 0x10);
        else
            return SetControl(0x10, 0x00);
    }

    public bool GetOutput() {
        return GetControl()["output"];
    }

    public bool SetOutput(bool output) {
        if (output)
            return SetControl(0x01, 0x01); //on
        else
            return SetControl(0x01, 0x00); //off
    }

    public Dictionary<string, bool> GetActual() {
        var actual = new Dictionary<string, bool>();
        byte[] ans = GetBinary(71);
        actual["remote"] = Convert.ToBoolean(ans[0] & 0x03) ? true : false;
        actual["local"] = Convert.ToBoolean(ans[0] & 0x03) ? true : false;
        actual["on"] = Convert.ToBoolean(ans[0] & 0x03) ? true : false;
        actual["CC"] = Convert.ToBoolean(ans[0] & 0x03) ? true : false;
        actual["CV"] = Convert.ToBoolean(ans[0] & 0x03) ? true : false;
        actual["OVP"] = Convert.ToBoolean(ans[0] & 0x03) ? true : false;
        actual["OCP"] = Convert.ToBoolean(ans[0] & 0x03) ? true : false;
        actual["OPP"] = Convert.ToBoolean(ans[0] & 0x03) ? true : false;
        actual["OTP"] = Convert.ToBoolean(ans[0] & 0x03) ? true : false;
        return actual;
    }

    public Dictionary<string, float> GetActualVoltageCurrent() {
        var actual = new Dictionary<string, float>();
        byte[] ans = GetBinary(71);
        actual["v"] = u_nom * ((ans[2] << 8) + ans[3]) / 25600;
        actual["i"] = i_nom * ((ans[4] << 8) + ans[5]) / 25600;
        return actual;
    }

    public float GetActualVoltage() {
        return GetActualVoltageCurrent()["v"];
    }

    public float GetActualCurrent() {
        return GetActualVoltageCurrent()["i"];
    }
}
}
