namespace DesktopSteps;

internal sealed class ReadableButton : Button
{
    public ReadableButton() => FlatStyle = FlatStyle.Flat;

    protected override void OnPaint(PaintEventArgs e)
    {
        if (Enabled || BackColor.GetBrightness() >= 0.5f)
        {
            base.OnPaint(e);
            return;
        }

        e.Graphics.Clear(BackColor);
        ControlPaint.DrawBorder(e.Graphics, ClientRectangle, Color.FromArgb(100, 120, 145), ButtonBorderStyle.Solid);
        var textBounds = ClientRectangle;
        if (Image is not null)
        {
            var imageX = 8;
            var imageY = Math.Max(0, (Height - Image.Height) / 2);
            e.Graphics.DrawImage(Image, imageX, imageY, Image.Width, Image.Height);
            textBounds = new Rectangle(imageX + Image.Width + 4, 0,
                Math.Max(0, Width - imageX - Image.Width - 8), Height);
        }
        TextRenderer.DrawText(e.Graphics, Text, Font, textBounds, Color.LightSkyBlue,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }
}
