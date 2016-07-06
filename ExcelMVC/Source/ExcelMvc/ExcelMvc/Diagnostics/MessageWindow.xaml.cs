namespace ExcelMvc.Diagnostics
{
    using System;
    using System.ComponentModel;
    using System.Runtime.CompilerServices;
    using System.Windows;
    using System.Windows.Input;
    using Runtime;

    /// <summary>
    /// Implements a visual sink for exception and information messages
    /// </summary>
    public partial class MessageWindow
    {
        #region Constructors
        private MessageWindow()
        {
            InitializeComponent();
            Closed += MessageWindow_Closed;
            Closing += MessageWindow_Closing;
        }

        #endregion

        #region Properties
        private static MessageWindow Instance { get; set; }

        private Message Model
        {
            get { return (Message)LayoutRoot.DataContext; }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Creates and shows to the status window
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void ShowInstance()
        {
            AsyncActions.Post(
                state =>
                {
                    if (Instance == null)
                        Instance = new MessageWindow();

                    // var interop = new WindowInteropHelper(Instance) { Owner = App.Instance.MainWindow.Handle };
                    Instance.Show();
                },
                null,
                false);
        }

        /// <summary>
        /// Hides the singleton
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void HideInstance()
        {
            AsyncActions.Post(
                state =>
                {
                    if (Instance != null)
                        Instance.Hide();
                },
                null,
                false);
        }

        /// <summary>
        /// Adds an exception to the status window
        /// </summary>
        /// <param name="ex">Exception to be addded</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void AddErrorLine(Exception ex)
        {
            AsyncActions.Post(
                state =>
                {
                    CreateInstance();
                    Instance.Model.AddErrorLine((Exception)state);
                }, 
                ex,
                false);
        }

        /// <summary>
        /// Adds an error to  to the status window
        /// </summary>
        /// <param name="error">Error to be added</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void AddErrorLine(string error)
        {
            AsyncActions.Post(
                state =>
                {
                    CreateInstance();
                    Instance.Model.AddErrorLine((string)state);
                },
                error,
                false);
        }

        /// <summary>
        /// Adds a message to  to the status window
        /// </summary>
        /// <param name="message">Message to be added</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void AddInfoLine(string message)
        {
            AsyncActions.Post(
                state =>
                {
                    CreateInstance();
                    Instance.Model.AddInfoLine((string)state);
                }, 
                message,
                false);
        }

        private static void CreateInstance()
        {
            if (Instance == null)
                Instance = new MessageWindow();
        }

        private static void MessageWindow_Closing(object sender, CancelEventArgs e)
        {
            if (ReferenceEquals(sender, Instance))
            {
                e.Cancel = true;
                HideInstance();
            }
        }

        private static void MessageWindow_Closed(object sender, EventArgs e)
        {
            if (ReferenceEquals(sender, Instance))
                Instance = null;
        }
        #endregion

        private void ButtonClear_OnClick(object sender, RoutedEventArgs e)
        {
            Model.Clear();
        }

        private void ButtonHide_OnClick(object sender, RoutedEventArgs e)
        {
            Hide();
        }

        private void LineLimit_OnKeyDown(object sender, KeyEventArgs e)
        {
            var key = Convert.ToInt32(e.Key);
            e.Handled = (key < Convert.ToInt32(Key.D0) || key > Convert.ToInt32(Key.D9))
                     && (key < Convert.ToInt32(Key.NumPad0) || key > Convert.ToInt32(Key.NumPad9));
        }
    }
}
