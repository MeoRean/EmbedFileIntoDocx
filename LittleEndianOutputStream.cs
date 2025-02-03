namespace EmbedFileIntoDocx;

public class LittleEndianOutputStream :IDisposable
{
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (null != out1)
            {
                out1.Dispose();
                out1 = null;
            }
        }
    }

    public void Close()
    {
        Dispose();
    }

    protected internal Stream out1 = null;

    public LittleEndianOutputStream(Stream out1)
    {
        this.out1 = out1;
    }

    public void WriteByte(int v)
    {
        out1.WriteByte((byte)v);
    }

    public void WriteDouble(double v)
    {
        WriteLong(BitConverter.DoubleToInt64Bits(v));
    }

    public void WriteInt(int v)
    {
        int b3 = (v >> 24) & 0xFF;
        int b2 = (v >> 16) & 0xFF;
        int b1 = (v >> 8) & 0xFF;
        int b0 = (v >> 0) & 0xFF;
        out1.WriteByte((byte)b0);
        out1.WriteByte((byte)b1);
        out1.WriteByte((byte)b2);
        out1.WriteByte((byte)b3);
    }

    public void WriteLong(long v)
    {
        WriteInt((int)(v >> 0));
        WriteInt((int)(v >> 32));
    }

    public void WriteShort(int v)
    {
        int b1 = (v >> 8) & 0xFF;
        int b0 = (v >> 0) & 0xFF;

        out1.WriteByte((byte)b0);
        out1.WriteByte((byte)b1);
    }

    public void Write(byte[] b)
    {
        out1.Write(b, 0, b.Length);
    }

    public void Write(byte[] b, int off, int len)
    {

        out1.Write(b, off, len);
    }

    public void Flush()
    {
        out1.Flush();
    }
}