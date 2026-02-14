using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Rendering;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
using ICSharpCode.AvalonEdit;
using System.Windows;

public class TextMarkerService : DocumentColorizingTransformer, IBackgroundRenderer
{
    private readonly TextSegmentCollection<TextMarker> markers;
    private readonly TextEditor editor;

    public TextMarkerService(TextEditor editor)
    {
        this.editor = editor;
        markers = new TextSegmentCollection<TextMarker>(editor.Document);

        editor.TextArea.TextView.BackgroundRenderers.Add(this);
        editor.TextArea.TextView.LineTransformers.Add(this);
    }

    public KnownLayer Layer => KnownLayer.Selection;

    public TextMarker Create(int startOffset, int length)
    {
        var marker = new TextMarker(startOffset, length);
        markers.Add(marker);
        editor.TextArea.TextView.Redraw();
        return marker;
    }

    public void RemoveAll()
    {
        markers.Clear();
        editor.TextArea.TextView.Redraw();
    }

    public void Draw(TextView textView, DrawingContext drawingContext)
    {
        if (markers == null || !markers.Any())
            return;

        foreach (var marker in markers)
        {
            foreach (var rect in BackgroundGeometryBuilder.GetRectsForSegment(textView, marker))
            {
                var brush = new SolidColorBrush(marker.BackgroundColor);
                drawingContext.DrawRectangle(brush, null, rect);
            }
        }
    }

    protected override void ColorizeLine(DocumentLine line)
    {
        foreach (var marker in markers.FindOverlappingSegments(line))
        {
            ChangeLinePart(
                marker.StartOffset,
                marker.EndOffset,
                element =>
                {
                    if (marker.ForegroundColor.HasValue)
                        element.TextRunProperties.SetForegroundBrush(
                            new SolidColorBrush(marker.ForegroundColor.Value));
                });
        }
    }

    public class TextMarker : TextSegment
    {
        public TextMarker(int startOffset, int length)
        {
            StartOffset = startOffset;
            Length = length;
        }

        public Color BackgroundColor { get; set; } = Colors.Yellow;
        public Color? ForegroundColor { get; set; }
    }
}
