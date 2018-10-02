using Spire.Presentation;
using Spire.License;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;

namespace Console.Spire
{
	class Program
	{
		static void Main(string[] args)
		{
			System.Console.WriteLine("testing...");
			using (var presentation = new Presentation())
			{
				presentation.LoadFromFile("test.pptx");

				var desiredX = 1200;
				var desiredY = 800;

				var scaleX = presentation.SlideSize.Size.Width / desiredX;
				var scaleY = presentation.SlideSize.Size.Height / desiredY;
				var ratio = Math.Max(scaleX, scaleY);

				var width = presentation.SlideSize.Size.Width / ratio;
				var height = presentation.SlideSize.Size.Height / ratio;

				List<Stream> result = new List<Stream>();
				for (int i = 0; i < presentation.Slides.Count; i++)
				{
					var slide = presentation.Slides[i];
					using (var image = slide.SaveAsImage((int)width, (int)height))
					{
						var stream = new FileStream(Path.GetTempFileName(), FileMode.Open, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose);
						image.Save("saved"+i+".jpeg", ImageFormat.Jpeg);
						stream.Seek(0, SeekOrigin.Begin);
					}
				}
			}
			System.Console.WriteLine("done");
			System.Console.Read();
		}
	}
}
