#if USE_IMAGESHARP
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Gif;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Drawing.Processing;
using SixLabors.Fonts;

namespace GrandBattleSupport
{
    internal static class ImageSharpGifExporter
    {
        // guildPtsPerFrame: optional list, where each element is a list of strings describing guild PTs for that frame/day
        public static void Export(List<Bitmap> bitmaps, string path, List<List<string>>? guildPtsPerFrame = null)
        {
            if (bitmaps == null || bitmaps.Count == 0) throw new ArgumentException("bitmaps required");

            List<SixLabors.ImageSharp.Image<Rgba32>> frames = new();
            SixLabors.ImageSharp.Image<Rgba32>? outImg = null;
            try
            {
                // Convert bitmaps to ImageSharp images. Optionally draw guild PTs onto each bitmap before conversion.
                for (int i = 0; i < bitmaps.Count; i++)
                {
                    var bmp = bitmaps[i];

                    var arr = BitmapToRgba32(bmp);
                    var img = SixLabors.ImageSharp.Image.LoadPixelData<Rgba32>(arr, bmp.Width, bmp.Height);

                    // If guild PTs provided for this frame, draw them using ImageSharp (cross-platform)
                    if (guildPtsPerFrame != null && i < guildPtsPerFrame.Count)
                    {
                        DrawGuildPtsOnImage(img, guildPtsPerFrame[i]);
                    }

                    frames.Add(img);
                }

                // Compose GIF: clone the first image and append clones of other frames
                var first = frames[0];
                outImg = first.Clone();

                for (int i = 1; i < frames.Count; i++)
                {
                    var f = frames[i];
                    // Add the frame's RootFrame directly
                    outImg.Frames.AddFrame(f.Frames.RootFrame);
                }

                // Set frame delay for all frames (100 = 1 second)
                try
                {
                    foreach (var frame in outImg.Frames)
                    {
                        frame.Metadata.GetGifMetadata().FrameDelay = 100;
                    }
                }
                catch
                {
                    // If metadata API unavailable for some versions, ignore and continue
                }

                var encoder = new GifEncoder();
                using var fs = File.Open(path, FileMode.Create, FileAccess.Write);
                outImg.Save(fs, encoder);
            }
            catch (Exception ex)
            {
                try
                {
                    File.WriteAllText("gif_error.log", ex.ToString());
                }
                catch { }
                throw;
            }
            finally
            {
                // Dispose images created
                if (outImg is not null) outImg.Dispose();
                foreach (var f in frames) f.Dispose();
            }
        }

        private static Rgba32[] BitmapToRgba32(Bitmap bmp)
        {
            var w = bmp.Width;
            var h = bmp.Height;
            var arr = new Rgba32[w * h];

            var rect = new System.Drawing.Rectangle(0, 0, w, h);
            var data = bmp.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            try
            {
                var bytes = new byte[Math.Abs(data.Stride) * h];
                System.Runtime.InteropServices.Marshal.Copy(data.Scan0, bytes, 0, bytes.Length);
                int stride = data.Stride;
                for (int y = 0; y < h; y++)
                {
                    for (int x = 0; x < w; x++)
                    {
                        int idx = y * stride + x * 4;
                        byte b = bytes[idx + 0];
                        byte g = bytes[idx + 1];
                        byte r = bytes[idx + 2];
                        byte a = bytes[idx + 3];
                        arr[y * w + x] = new Rgba32(r, g, b, a);
                    }
                }
            }
            finally
            {
                bmp.UnlockBits(data);
            }

            return arr;
        }

        private static void DrawGuildPtsOnImage(SixLabors.ImageSharp.Image<Rgba32> img, List<string>? guildPts)
        {
            if (guildPts == null || guildPts.Count == 0) return;

            int w = img.Width;

            // Choose font size relative to image width
            float fontSize = Math.Max(12f, w / 40f);
            var font = SixLabors.Fonts.SystemFonts.CreateFont("Segoe UI", fontSize, SixLabors.Fonts.FontStyle.Bold);

            // Approximate line height and width (avoid TextMeasurer to keep dependency surface small)
            float lineHeight = fontSize * 1.2f;
            float padding = Math.Max(6f, fontSize / 4f);
            float totalHeight = lineHeight * guildPts.Count + padding * 2f;

            var bgColor = new Rgba32(0, 0, 0, 180);
            var textColor = new Rgba32(255, 255, 255, 255);
            var shadowColor = new Rgba32(0, 0, 0, 160);

            img.Mutate(ctx =>
            {
                // Fill background rectangle
                ctx.Fill(bgColor, new SixLabors.ImageSharp.RectangleF(0, 0, w, totalHeight));

                for (int i = 0; i < guildPts.Count; i++)
                {
                    var text = guildPts[i] ?? string.Empty;
                    // approximate width by character count
                    float approxWidth = text.Length * fontSize * 0.6f;
                    float x = (w - approxWidth) / 2f;
                    float y = padding + i * lineHeight;

                    // Draw shadow then text
                    ctx.DrawText(text, font, shadowColor, new SixLabors.ImageSharp.PointF(x + 1, y + 1));
                    ctx.DrawText(text, font, textColor, new SixLabors.ImageSharp.PointF(x, y));
                }
            });
        }
    }
}
#endif