using System.Collections.Generic;

public class InfillType
{
    public string ConfigName { get; }
    public string DescPrefix { get; }
    public string LenPrefix { get; }
    public int Offset { get; } = 0;

    public InfillType(string configName, string descPrefix, string lenPrefix, int offset)
    {
        ConfigName = configName;
        DescPrefix = descPrefix;
        LenPrefix = lenPrefix;
        Offset = offset;
    }
}

