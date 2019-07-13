using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using log4net.Appender;
using log4net.Core;
using log4net.Layout;
namespace InsuranceCompareTool.ShareCommon
{
    /// <summary>
    /// Visual Logger
    /// </summary>
    public sealed class VisualLogger : AppenderSkeleton
    {
        private static readonly Dictionary<Level, Brush> COLORS = new Dictionary<Level, Brush>();
        private readonly PatternLayout mLayout;
        private RichTextBox mTextBox;
        static VisualLogger()
        {
            COLORS.Add(Level.Debug, Brushes.Blue);
            COLORS.Add(Level.Error, Brushes.Red);
            COLORS.Add(Level.Info, Brushes.Black);
            COLORS.Add(Level.Warn, Brushes.Purple);
        }
        public VisualLogger()
        {
            mLayout = new PatternLayout();
            mLayout.ConversionPattern = "%d{HH:mm:ss} %m";
            mLayout.ActivateOptions();
            Layout = mLayout;
            ActivateOptions();
        }
        /// <summary>
        /// TextBox control for shows event logs.
        /// </summary>
        public RichTextBox TextBox
        {
            get
            {
                return mTextBox;
            }
            set
            {
                mTextBox = value;
                if (value != null)
                {
                    mTextBox.Document.LineHeight = 1;
                }
            }
        }
        protected override void Append(LoggingEvent loggingEvent)
        {
            if (TextBox == null)
            {
                return;
            }
            TextBox.Dispatcher.Invoke(() =>
            {
                TextPointer position = TextBox.CaretPosition.DocumentStart;
                var lineBreak = new LineBreak(position);
                var text = new Run(mLayout.Format(loggingEvent), position);
                text.Foreground = COLORS[loggingEvent.Level];
            });
        }
 
    }
}
