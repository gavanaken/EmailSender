using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace EmailSender
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        PreviewWindow previewWin;
        public MainWindow()
        {
            InitializeComponent();
            Email_Box.Text = Properties.Resources.Default_Email;
            Substitution_String.Text = Properties.Resources.Default_Substitution_String;
            Subject_Box.Text = Properties.Resources.Default_Subject;
            InitializeGroupList();
            InitializeCanvas();
        }

        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate { }));
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void From_CSV_Button_Click(object sender, RoutedEventArgs e)
        {

            lFnLoadFileData();
        }

        private void Send_Test_Button_Click(object sender, RoutedEventArgs e)
        {
            (int code, string response) = Send_Email(Email_Box.Text, GetSubstitutedBody(Properties.Resources.Default_User_Name));
            StatusBarMessage.Text = response;
        }

        private void Send_Emails_Button_Click(object sender, RoutedEventArgs e)
        {
            int numGroups = Group_List.Items.Count;
            MessageBoxResult messageBoxResult = MessageBox.Show("You are about to send whatever is in the " +
                "canvas to " + (numGroups - 1).ToString() + " recipient(s). There is no going back! " +
                "Are you sure you want to do this?", "Send Confirmation", MessageBoxButton.YesNo);
            if (messageBoxResult == MessageBoxResult.Yes)
            {
                Progress_Bar.Visibility = Visibility.Visible;
                Progress_Bar.Maximum = numGroups - 1;
                Progress_Bar.Value = 0;
                for (int i = 0; i < numGroups; i++)
                {
                    DataGridRow row = (DataGridRow)Group_List.ItemContainerGenerator.ContainerFromIndex(i);
                    row.IsSelected = true;
                    if (row.Item.GetType() != typeof(EmailSender.Group))
                        continue;
                    Group grp = (Group)row.Item;
                    (int code, string response) = Send_Email(grp.email, GetSubstitutedBody(grp.name));
                    StatusBarMessage.Text = response;
                    switch (code)
                    {
                        case 0:
                            break;
                        case 1:
                            // Authentication error, stop trying to send emails
                            Progress_Bar.Visibility = Visibility.Hidden;
                            return;
                        case -1:
                            // Unsure what happened, but keep trying
                            break;
                    }
                    DoEvents();
                    Progress_Bar.Value++;
                    DoEvents();
                }
            }
        }

        private void Preview_Button_Click(object sender, RoutedEventArgs e)
        {
            Group selectedGroup;
            if (Group_List.SelectedItems.Count == 0)
            {
                selectedGroup = (Group) Group_List.Items[0];
            }
            else
            {
                try
                {
                    selectedGroup = (Group)Group_List.SelectedItems[0];
                }
                catch (Exception exc)
                {
                    return;
                }

            }
            DisplayPreview(selectedGroup);
        }

        private string GetSubstitutedHtml(string substitution)
        {
            return Html_Editor.DocumentHtml.Replace(Substitution_String.Text, substitution);
        }

        private string GetSubstitutedBody(string substitution)
        {
            return Html_Editor.BodyHtml.Replace(Substitution_String.Text, substitution);
        }

        void lFnLoadFileData()
        {
            Microsoft.Win32.OpenFileDialog lObjFileDlge = new Microsoft.Win32.OpenFileDialog();
            lObjFileDlge.Filter = "CSV Files|*.csv";
            lObjFileDlge.FilterIndex = 1;
            lObjFileDlge.Multiselect = false;
            string fName = "";
            bool? lBlnUserclicked = lObjFileDlge.ShowDialog();
            if (lBlnUserclicked != null || lBlnUserclicked == true)
            {
                fName = lObjFileDlge.FileName;
            }
            if (System.IO.File.Exists(fName) == true)
            {
                StreamReader lObjStreamReader = new StreamReader(fName);
                lFnGenerateData(lObjStreamReader);
                lObjStreamReader.Close();
            }
        }

        void lFnGenerateData(StreamReader aReader)
        {
            List<Group> lstGroupList = new List<Group>();
            Group_List.Columns.Clear();
            while (aReader.Peek() > 0)
            {
                string lStrLine = aReader.ReadLine();
                if (lStrLine == null)
                    break;
                if (lStrLine.Trim() == "")
                    continue;
                string[] lArrStrCells = null;
                lArrStrCells = lStrLine.Split(';');
                if (lArrStrCells == null)
                    continue;

                Group groupInfo = new Group();
                if (lArrStrCells.Length != 3)
                {
                    break;
                }
                groupInfo.name = lArrStrCells[0];
                groupInfo.email = lArrStrCells[1];
                groupInfo.school = lArrStrCells[2];

                lstGroupList.Add(groupInfo);
            }
            aReader.Close();
            Group_List.ItemsSource = lstGroupList;

        }

        private void InitializeGroupList()
        {
            Group_List.Columns.Clear();
            List<Group> lstGroupList = new List<Group>();
            Group nullGroup = new Group();
            nullGroup.name = Properties.Resources.Default_User_Name;
            nullGroup.email = Properties.Resources.Default_Email;
            nullGroup.school = Properties.Resources.Default_School;
            lstGroupList.Add(nullGroup);
            Group_List.ItemsSource = lstGroupList;
        }

        private void InitializeCanvas()
        {
            string example_text = ReadTextResourceFromAssembly("EmailSender.static.Example.html");
            Html_Editor.BodyHtml = example_text;
        }

        private void InitializePreviewWindow()
        {
            if (this.previewWin == null)
            {
                this.previewWin = new PreviewWindow();
                this.previewWin.Closed += (sender, args) => this.previewWin = null;
                this.previewWin.Show();
            }
        }

        private void DisplayPreview(Group selectedGroup)
        {
            try
            {
                InitializePreviewWindow();
                previewWin.Display(GetSubstitutedHtml(selectedGroup.name));
                StatusBarMessage.Text = String.Format("Displaying preview for {0}", selectedGroup.name);
            }
            catch (Exception exc)
            {
                StatusBarMessage.Text = "Error loading preview";
            }
        }

        private void DisplayHelp()
        {
            try
            {
                InitializePreviewWindow();
                previewWin.Display(ReadTextResourceFromAssembly("EmailSender.static.Example.html"));
            }
            catch (Exception exc)
            {
                StatusBarMessage.Text = "Error loading help";
            }
        }

        private void Group_List_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Row_KeyPress(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                Group selectedGroup = (Group)((DataGridRow)sender).Item;
                DisplayPreview(selectedGroup);
            }
        }

        bool IsValidEmail(string email)
        {
            try
            {
                var addr = new MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        private static string ReadTextResourceFromAssembly(string name)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                return new StreamReader(stream).ReadToEnd();
            }
        }

        private bool Validate_Email(string recipient, string message)
        {
            if(string.IsNullOrWhiteSpace(Email_Box.Text))
            {
                StatusBarMessage.Text = "Email cannot be empty";
                return false;
            }
            if (string.IsNullOrWhiteSpace(Password_Box.Password))
            {
                StatusBarMessage.Text = "Password cannot be empty";
                return false;
            }
            if (string.IsNullOrWhiteSpace(Substitution_String.Text))
            {
                StatusBarMessage.Text = "Substitution string cannot be empty";
                return false;
            }
            if (!IsValidEmail(recipient))
            {
                return false;
            }
            if (!IsValidEmail(Email_Box.Text))
            {
                return false;
            }
            return true;
        }

        private (int, string) Send_Email(string recipient, string message)
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(Email_Box.Text);
            mail.To.Add(recipient);
            mail.Subject = Subject_Box.Text;
            mail.Body = message;
            mail.IsBodyHtml = true;

            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential(Email_Box.Text, Password_Box.Password);
            SmtpServer.EnableSsl = true;
            try
            {
                SmtpServer.Send(mail);
            }
            catch (SmtpException e)
            {
                if (e.Message.Contains("Authentication"))
                    return (1, "Error: Email/password incorrect or less secure apps OFF");
                return (-1, "Failure: " + String.Copy(recipient));
            }

            return (0, "Success: " + String.Copy(recipient));
        }

        private void Help_Button_Click(object sender, RoutedEventArgs e)
        {
            DisplayHelp();
        }
    }
}
