﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OpenXmlPowerTools;

namespace ExamplePresentatonBuilder01
{
    class ExamplePresentationBuilder01
    {
        static void Main(string[] args)
        {
            string source1 = "../../Contoso.pptx";
            string source2 = "../../Companies.pptx";
            string source3 = "../../Customer Content.pptx";
            string source4 = "../../Presentation One.pptx";
            string source5 = "../../Presentation Two.pptx";
            string source6 = "../../Presentation Three.pptx";
            string contoso1 = "../../Contoso One.pptx";
            string contoso2 = "../../Contoso Two.pptx";
            string contoso3 = "../../Contoso Three.pptx";
            List<SlideSource> sources = null;

            var sourceDoc = new PmlDocument(source1);
            sources = new List<SlideSource>()
            {
                new SlideSource(sourceDoc, 0, 1, false),  // Title
                new SlideSource(sourceDoc, 1, 1, false),  // First intro (of 3)
                new SlideSource(sourceDoc, 4, 2, false),  // Sales bios
                new SlideSource(sourceDoc, 9, 3, false),  // Content slides
                new SlideSource(sourceDoc, 13, 1, false),  // Closing summary
            };
            PresentationBuilder.BuildPresentation(sources, "Out1.pptx");

            sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source2), 2, 1, true),  // Choose company
                new SlideSource(new PmlDocument(source3), false),       // Content
            };
            PresentationBuilder.BuildPresentation(sources, "Out2.pptx");

            sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(source4), true),
                new SlideSource(new PmlDocument(source5), true),
                new SlideSource(new PmlDocument(source6), true),
            };
            PresentationBuilder.BuildPresentation(sources, "Out3.pptx");

            sources = new List<SlideSource>()
            {
                new SlideSource(new PmlDocument(contoso1), true),
                new SlideSource(new PmlDocument(contoso2), true),
                new SlideSource(new PmlDocument(contoso3), true),
            };
            PresentationBuilder.BuildPresentation(sources, "Out4.pptx");
        }
    }
}
