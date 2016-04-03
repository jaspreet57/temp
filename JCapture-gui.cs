using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace JCaptureApp {
    public class WinFormExample : Form {

        private Button fileSelector;
		private Button about;
		private Button capture;
		private string wordDocPath;
		private ScreenCapture sc;
		private Image img;
		private Microsoft.Office.Interop.Word.Application ap;
		private Document doc;
		private Selection sel;

        public WinFormExample() {
			sc = new ScreenCapture();
            DisplayGUI();
        }

        private void DisplayGUI() {
            this.Name = "JCapture-snip2doc";
            this.Text = "JCapture-snip2doc";
			this.TopMost = true;
            this.Size = new Size(300, 110);
            this.StartPosition = FormStartPosition.CenterScreen;

            fileSelector = new Button();
            fileSelector.Name = "fileSelector";
            fileSelector.Text = "Open Document";
            fileSelector.Size = new Size(this.Width - 200, this.Height - 60);
            fileSelector.Location = new System.Drawing.Point(
                (this.Width - fileSelector.Width) / 8 ,
                (this.Height - fileSelector.Height) / 6);
            fileSelector.Click += new System.EventHandler(this.FileSelectorClick);

            this.Controls.Add(fileSelector);
			
			
			
			about = new Button();
            about.Name = "about";
            about.Text = "About JCapture";
            about.Size = new Size(this.Width - 200, this.Height - 60);
            about.Location = new System.Drawing.Point(
                (this.Width - about.Width - 50) ,
                (this.Height - about.Height) / 6);
            about.Click += new System.EventHandler(this.AboutClick);
            this.Controls.Add(about);
        }
		
		private void AboutClick(object source, EventArgs e) {
             MessageBox.Show("JCapture v1.0 \n JCapture is multipurpose snipping tool. \n Its current version is capable of taking screenshot and saving in ms office document with just one click. \n This tool can be used in tasks like testing and documenting. \n Designed and Developed by Jaspreet Singh \"961923\"");
        }
		
		
		private void FileSelectorClick(object sender, EventArgs e)
		{
			OpenFileDialog im = new OpenFileDialog();
            if (im.ShowDialog() == DialogResult.OK)
            {
				wordDocPath = im.FileName;
				Console.WriteLine("File Selected is " + wordDocPath);
				
				//open word file here
				ap = new Microsoft.Office.Interop.Word.Application();
				ap.Application.Visible = true;	
				try
				{
					doc = ap.Documents.Open( @wordDocPath, ReadOnly: false, Visible: true );
					doc.Activate();
					ap.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize;
					
					MessageBox.Show("Start taking screenshots by clicking on Capture button.");
				
					this.Controls.Remove(fileSelector);
					fileSelector.Dispose();
					capture = new Button();
					capture.Name = "captureButton";
					capture.Text = "Capture";
					capture.Size = new Size(this.Width - 200, this.Height - 60);
					capture.Location = new System.Drawing.Point(
						(this.Width - capture.Width) / 8,
						(this.Height - capture.Height) / 6);
					capture.Click += new System.EventHandler(this.CaptureClick);

					this.Controls.Add(capture);
				}catch ( Exception ex )
				{
					MessageBox.Show("Unable to open word document. Refer Console for proper error message.");
					Console.WriteLine( "Exception Caught: " + ex.Message ); // Could be that the document is already open (/) or Word is in Memory(?)
					// Ambiguity between method 'Microsoft.Office.Interop.Word._Application.Quit(ref object, ref object, ref object)' and non-method 'Microsoft.Office.Interop.Word.ApplicationEvents4_Event.Quit'. Using method group.
					// ap.Quit( SaveChanges: false, OriginalFormat: false, RouteDocument: false );
					try{
						( (_Application)ap ).Quit( SaveChanges: false, OriginalFormat: false, RouteDocument: false );
						System.Runtime.InteropServices.Marshal.ReleaseComObject( ap );
					}catch (Exception exception){
						Console.WriteLine(exception.ToString());
					}
					
				}
            }else{
				MessageBox.Show("Please select any word document to save screenshots");
			}
		}
		
		private void CaptureClick(object source, EventArgs e) {
			img = sc.CaptureScreen();
			Clipboard.SetImage(img);
			try {
				if(ap != null){
					sel = null;
					sel =  ap.Selection;
			 
					if ( sel != null )
					{
						switch ( sel.Type )
						{
							case WdSelectionType.wdSelectionIP:
								
								sel.Paste();
								sel.TypeText("");
								sel.TypeParagraph();
								sel.TypeText("");
								sel.TypeParagraph();
								break;
				 
							default:
								MessageBox.Show("Some problem is with your document. I am unable to write !");
								break;
				 
						}
				 
						ap.Documents.Save( NoPrompt: true, OriginalFormat: true );
				 
					}
					else
					{
						MessageBox.Show("Some problem is with your document. I am unable to write !");
					}
				}
			}catch(Exception ex){
				Console.WriteLine(ex.ToString());
				MessageBox.Show("Now Where should I save screenshot? You have closed document yourself !... dont do it next time");
				this.Close();
			}
			
        }
		
		protected override void OnFormClosing(FormClosingEventArgs e)
		{
			if(ap != null){
				try{
					
					ap.Documents.Close( SaveChanges: false, OriginalFormat: false, RouteDocument: false );
				}catch ( Exception ex )
				{
					Console.WriteLine(ex.ToString());
					MessageBox.Show("Document is either already closed or not working!");
				}
			   // Ambiguity between method 'Microsoft.Office.Interop.Word._Application.Quit(ref object, ref object, ref object)' and non-method 'Microsoft.Office.Interop.Word.ApplicationEvents4_Event.Quit'. Using method group.
				// ap.Quit( SaveChanges: false, OriginalFormat: false, RouteDocument: false );
				try{
					( (_Application)ap ).Quit( SaveChanges: false, OriginalFormat: false, RouteDocument: false );
					System.Runtime.InteropServices.Marshal.ReleaseComObject( ap );
				}catch(Exception ex){
					Console.WriteLine(ex.ToString());
				}
			}
		   
		    base.OnFormClosing(e);
		}   
		
		
		[STAThread]
        public static void Main(String[] args) {
            System.Windows.Forms.Application.Run(new WinFormExample());
        }
		
		
		private class ScreenCapture
		{
			/// <summary>
			/// Creates an Image object containing a screen shot of the entire desktop
			/// </summary>
			/// <returns></returns>
			public Image CaptureScreen()
			{
				return CaptureWindow( User32.GetDesktopWindow() );
			}
		
			public Image CaptureWindow(IntPtr handle)
			{
				// get te hDC of the target window
				IntPtr hdcSrc = User32.GetWindowDC(handle);
				// get the size
				User32.RECT windowRect = new User32.RECT();
				User32.GetWindowRect(handle,ref windowRect);
				int width = windowRect.right - windowRect.left;
				int height = windowRect.bottom - windowRect.top;
				// create a device context we can copy to
				IntPtr hdcDest = GDI32.CreateCompatibleDC(hdcSrc);
				// create a bitmap we can copy it to,
				// using GetDeviceCaps to get the width/height
				IntPtr hBitmap = GDI32.CreateCompatibleBitmap(hdcSrc,width,height);
				// select the bitmap object
				IntPtr hOld = GDI32.SelectObject(hdcDest,hBitmap);
				// bitblt over
				GDI32.BitBlt(hdcDest,0,0,width,height,hdcSrc,0,0,GDI32.SRCCOPY);
				// restore selection
				GDI32.SelectObject(hdcDest,hOld);
				// clean up
				GDI32.DeleteDC(hdcDest);
				User32.ReleaseDC(handle,hdcSrc);
				// get a .NET image object for it
				Image img = Image.FromHbitmap(hBitmap);
				// free up the Bitmap object
				GDI32.DeleteObject(hBitmap);
				return img;
			}

			private class GDI32
			{

				public const int SRCCOPY = 0x00CC0020; // BitBlt dwRop parameter
				[DllImport("gdi32.dll")]
				public static extern bool BitBlt(IntPtr hObject,int nXDest,int nYDest,
					int nWidth,int nHeight,IntPtr hObjectSource,
					int nXSrc,int nYSrc,int dwRop);
				[DllImport("gdi32.dll")]
				public static extern IntPtr CreateCompatibleBitmap(IntPtr hDC,int nWidth,
					int nHeight);
				[DllImport("gdi32.dll")]
				public static extern IntPtr CreateCompatibleDC(IntPtr hDC);
				[DllImport("gdi32.dll")]
				public static extern bool DeleteDC(IntPtr hDC);
				[DllImport("gdi32.dll")]
				public static extern bool DeleteObject(IntPtr hObject);
				[DllImport("gdi32.dll")]
				public static extern IntPtr SelectObject(IntPtr hDC,IntPtr hObject);
			}

			
			private class User32
			{
				[StructLayout(LayoutKind.Sequential)]
				public struct RECT
				{
					public int left;
					public int top;
					public int right;
					public int bottom;
				}
				[DllImport("user32.dll")]
				public static extern IntPtr GetDesktopWindow();
				[DllImport("user32.dll")]
				public static extern IntPtr GetWindowDC(IntPtr hWnd);
				[DllImport("user32.dll")]
				public static extern IntPtr ReleaseDC(IntPtr hWnd,IntPtr hDC);
				[DllImport("user32.dll")]
				public static extern IntPtr GetWindowRect(IntPtr hWnd,ref RECT rect);
			}
		}
    }
}
