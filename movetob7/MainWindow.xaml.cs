using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Keys = System.Windows.Forms.Keys;

using System.Collections;
using Shell32;
using System.IO;
using System.Threading;
using Newtonsoft.Json;

namespace movetob7
{
    public class ViewModelb7
    {
        private List<Directoryb7> _items;
        public List<Directoryb7> Items
        {
            get
            {
                return _items;
            }
            set
            {
                _items = value;
            }
        }
    }

    public class Directoryb7
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string FullPath { get; set; }
        public int Uses { get; set; }
    }


    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, IDisposable
    {

        private GlobalKeyboardHook _globalKeyboardHook;

         List<string> LastExplorerSelectedItems = null;

        ViewModelb7 viewModelOnScreen = null;

        public MainWindow()
        {
            InitializeComponent();

            SetupKeyboardHooks();

            LoadDirectories();


            this.PreviewKeyDown += new System.Windows.Input.KeyEventHandler(HandleWindowPreviewKeyDown);
        }


        private void HandleWindowPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                LastExplorerSelectedItems = null;
                Close();
            }
        }

        private void LoadDirectories()
        {
            var path = System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
            var json = System.IO.File.ReadAllText(path + "\\" + "directories.txt");

            try
            {
                var directories = JsonConvert.DeserializeObject<List<Directoryb7>>(json);

                viewModelOnScreen = new ViewModelb7();

                viewModelOnScreen.Items = directories;

                DataContext = viewModelOnScreen;
            }
            catch
            {
                System.Windows.MessageBox.Show("An error occured trying to read directories.txt");
            }
        }


        public void SetupKeyboardHooks()
        {
            _globalKeyboardHook = new GlobalKeyboardHook();
            _globalKeyboardHook.KeyboardPressed += OnKeyPressed;
        }

        private void OnKeyPressed(object sender, GlobalKeyboardHookEventArgs e)
        {
            //Debug.WriteLine(e.KeyboardData.VirtualCode);

            if (e.KeyboardData.VirtualCode != GlobalKeyboardHook.VkNUMPAD0)
                return;

            // seems, not needed in the life.
            //if (e.KeyboardState == GlobalKeyboardHook.KeyboardState.SysKeyDown &&
            //    e.KeyboardData.Flags == GlobalKeyboardHook.LlkhfAltdown)
            //{
            //    MessageBox.Show("Alt + Print Screen");
            //    e.Handled = true;
            //}
            //else

            if (e.KeyboardState == GlobalKeyboardHook.KeyboardState.KeyDown)
            {
                //System.Windows.MessageBox.Show("0 pressed");


                new Thread(new ThreadStart(() => {
                    LastExplorerSelectedItems = getExplorerSelectedItems();

                    System.Windows.Application.Current.Dispatcher.Invoke(new Action(() => {
                        this.Focus();
                        this.Activate();
                        inputTextBox.Focus();
                    }));
                })).Start();
                

                e.Handled = true;
            }
        }

        public void Dispose()
        {
            _globalKeyboardHook?.Dispose();
        }

        private List<string> getExplorerSelectedItems()
        {
            string filename;
            var selected = new List<string>();
            foreach (SHDocVw.InternetExplorer window in new SHDocVw.ShellWindowsClass())
            {
                filename = System.IO.Path.GetFileNameWithoutExtension(window.FullName).ToLower();
                if (filename.ToLowerInvariant() == "explorer")
                {
                    Shell32.FolderItems items = ((Shell32.IShellFolderViewDual2)window.Document).SelectedItems();
                    foreach (Shell32.FolderItem item in items)
                    {
                        selected.Add(item.Path);
                    }
                }
            }
            return selected;
        }

        private void inputTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                //find it and move it

                //get destionation
                string destionation = "";
                foreach(var directoryb7 in viewModelOnScreen.Items)
                {
                    if(inputTextBox.Text == directoryb7.Name)
                    {
                        destionation = directoryb7.FullPath;

                        if(System.IO.Directory.Exists(destionation) == false)
                        {
                            System.IO.Directory.CreateDirectory(destionation);
                        }
                        break;
                    }
                }
                if(destionation == "")
                {
                    System.Windows.MessageBox.Show("Not a right choice!");
                    return;
                }


                if(LastExplorerSelectedItems != null)
                {
                    foreach(var fileOrDirectory in LastExplorerSelectedItems)
                    {
                        try
                        {
                            if (System.IO.Directory.Exists(fileOrDirectory))
                            {
                                System.IO.Directory.Move(fileOrDirectory, destionation);
                            }
                            else
                            {
                                var filename = System.IO.Path.GetFileName(fileOrDirectory);
                                System.IO.File.Move(fileOrDirectory, destionation + filename);
                            }
                        }
                        catch
                        {
                            System.Windows.MessageBox.Show("An error occured");
                        }
                    }
                }
            }
        }
    }
}