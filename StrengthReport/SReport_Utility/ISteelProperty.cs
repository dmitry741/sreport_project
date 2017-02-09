namespace SReport_Utility
{
    interface ISteelProperty
    {
        double[] Alpha { get; set; }
        double[] E { get; set; }
        string Name { get; set; }
        double[] Rm { get; set; }
        double[] Rp { get; set; }
        double[] Sigma { get; set; }
    }
}
