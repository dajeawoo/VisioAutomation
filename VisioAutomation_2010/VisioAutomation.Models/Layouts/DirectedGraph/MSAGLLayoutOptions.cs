﻿namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class MsaglLayoutOptions : LayoutOptions
    {
        public double ScalingFactor { get; set; }
        public bool UseDynamicConnectors { get; set; }

        public MsaglLayoutOptions() :
            base()
        {
            this.UseDynamicConnectors = true;
            this.ScalingFactor = 14;
        }
    }
}