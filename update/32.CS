using System.Windows.Forms;
using System.Drawing;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Diagnostics; 
using System.ComponentModel;
using System.Text;
using System;
using System.Net;
using System.Collections.Generic; 

class updateform:Form
{
	string zipfile="",currentfile="";
	ProgressBar pbar; 
	ftp ftpClient;
    Label lbl1, lbl2, lbl3;
    Button btn1,btn2,btn3,btn4,btn5,btn7,btn8, btn9;
	
	public updateform()
	{
		ftpClient = new ftp(@"ftp://medionline.co.in/", "update@medionline.co.in", "123456");

        lbl1 = new Label{
            Text = "1.PLEASE ENSURE THAT YOU HAVE BACKUP OF SOFTWARE BEFORE PURSUING FOR UPDATE.\n2.PLEASE ENSURE THAT MEDI IS CLOSED(NOT RUNNING). ANY WHERE IN LAN ALSO.",
            Location = new Point(20, 30), AutoSize = true, Parent = this        };

        pbar = new ProgressBar
        {
            Maximum = 110,
            Location = new Point(20, 390),
            Width = 550,
            Height = 30,
            Parent = this
        };

        lbl2 = new Label
        {
            Location = new Point(20, 440),
            AutoSize = true,
            Parent = this
        };

        lbl3 = new Label
        {
            Location = new Point(400, 440),
            AutoSize = true,
            Parent = this
        };

        btn1 = new Button
        {
            Text = "CHECK FOR UPDATE !!",
            Width = 275,
            Location = new Point(20, 90),
            Visible = false,
            Parent = this
        };
        btn1.Click += new EventHandler(chkupd);

        btn2 = new Button
        {
            Text = "PRINTER: LASER, INVOICE SIZE:6 INCH, BUSINESS:PHARMA WHOLESALE",
            Width = 550,
            Location = new Point(20, 120),
            Parent = this
        };
        btn2.Click += new EventHandler(las6);

        btn3 = new Button
        {
            Text = "PRINTER: DOT MATRIX, INVOICE SIZE:6 INCH, BUSINESS:PHARMA WHOLESALE",
            Width = 550,
            Location = new Point(20, 150),
            Parent = this
        };
        btn3.Click += new EventHandler(dot6);

        btn4 = new Button
        {
            Text = "PRINTER: LASER, INVOICE SIZE:6 INCH AND 12 INCH (CHIOCE), BUSINESS:PHARMA WHOLESALE",
            Width = 550,
            Location = new Point(20, 180),
            Parent = this
        };
        btn4.Click += new EventHandler(las612);

        btn5 = new Button
        {
            Text = "PRINTER: LASER, INVOICE SIZE:12 INCH , BUSINESS:PHARMA WHOLESALE",
            Width = 550,
            Location = new Point(20, 210),
            Parent = this
        };
        btn5.Click += new EventHandler(las12);


        btn7 = new Button
        {
            Text = "PRINTER: DOT MATRIX, INVOICE SIZE:12 INCH, BUSINESS:PHARMA WHOLESALE",
            Width = 550,
            Location = new Point(20, 270),
            Parent = this
        };
        btn7.Click += new EventHandler(dot12);

        btn8 = new Button
        {
            Text = "PRINTER: LASER, INVOICE SIZE:6 INCH, BUSINESS:PHARMA RETAIL",
            Width = 550,
            Location = new Point(20, 300),
            Parent = this
        };
        btn8.Click += new EventHandler(ret);

        btn9 = new Button
        {
            Text = "UPDATE SOFTWARE USING CUSTOMIZE/FMCG INVOICE",
            Width = 550,
            Location = new Point(20, 360),
            Visible = false,
            Parent = this
        };
        btn9.Click += new EventHandler(cust);

        Size = new Size(640, 630);
        CenterToScreen();
    }
	void chkupd(object sender,EventArgs e)
	{}
	void las6(object sender,EventArgs e)
	{
		if(!File.Exists(Application.StartupPath+"\\MEDI.EXE")||!File.Exists(Application.StartupPath+"\\BILL.DBF"))
		{
			MessageBox.Show("KINDLY COPY AND PAST THIS PROGRAM TO YOUR MEDI FOLDER.\nUNABLE TO LOCATE MEDI APPLICATION HERE\nUNABLE TO GET THE FIRM NAME FOR UPDATE.");
			return;
		}
		string m = "YOU HAD CHOOSE TO UPDATE\nWITH PHARMACEUTICAL SOFTWARE!\nARE YOU SURE TO CONTINUE?\nPRESS YES TO CONTINUE OR PRESS NO TO CANCEL. ";
		DialogResult result = MessageBox.Show(this,m,"",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
		if(result ==DialogResult.Yes)
		{
			copy(Application.StartupPath,"..\\UPDATEBOTMEDIBACK"+DateTime.Now.ToString("ddMMyyyyhhmmss"));
			senddata();
			File.Delete(Application.StartupPath+"\\TOOLS.EXE");
            zipfile = "LASER.ZIP";
            downupd("las6");
			unzip();
		}
	}
	void dot6(object sender,EventArgs e)
	{
		if(!File.Exists(Application.StartupPath+"\\MEDI.EXE")||!File.Exists(Application.StartupPath+"\\BILL.DBF"))
		{
			MessageBox.Show("KINDLY COPY AND PAST THIS PROGRAM TO YOUR MEDI FOLDER.\nUNABLE TO LOCATE MEDI APPLICATION HERE\nUNABLE TO GET THE FIRM NAME FOR UPDATE.");
			return;
		}

		string m = "YOU HAD CHOOSE TO UPDATE\nWITH PHARMACEUTICAL SOFTWARE!\nARE YOU SURE TO CONTINUE?\nPRESS YES TO CONTINUE OR PRESS NO TO CANCEL. ";
		DialogResult result = MessageBox.Show(this,m,"",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
		if(result ==DialogResult.Yes)
		{
			copy(Application.StartupPath,"..\\UPDATEBOTMEDIBACK"+DateTime.Now.ToString("ddMMyyyyhhmmss"));
			senddata();		
			File.Delete(Application.StartupPath+"\\TOOLS.EXE");
            zipfile = "DOTHSN4.ZIP";
            downupd("dot6");
			unzip();
		}
	}
    void las612(object sender, EventArgs e)
    {
        if (!File.Exists(Application.StartupPath + "\\MEDI.EXE") || !File.Exists(Application.StartupPath + "\\BILL.DBF"))
        {
            MessageBox.Show("KINDLY COPY AND PAST THIS PROGRAM TO YOUR MEDI FOLDER.\nUNABLE TO LOCATE MEDI APPLICATION HERE\nUNABLE TO GET THE FIRM NAME FOR UPDATE.");
            return;
        }

        string m = "YOU HAD CHOOSE TO UPDATE\nWITH PHARMACEUTICAL SOFTWARE!\nARE YOU SURE TO CONTINUE?\nPRESS YES TO CONTINUE OR PRESS NO TO CANCEL. ";
        DialogResult result = MessageBox.Show(this, m, "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
        if (result == DialogResult.Yes)
        {
            copy(Application.StartupPath, "..\\UPDATEBOTMEDIBACK" + DateTime.Now.ToString("ddMMyyyyhhmmss"));
            senddata();
            File.Delete(Application.StartupPath + "\\TOOLS.EXE");
            zipfile = "OBJ612.ZIP";
            downupd("las612");
            unzip();
        }
    }
    void las12(object sender,EventArgs e)
	{
		if(!File.Exists(Application.StartupPath+"\\MEDI.EXE")||!File.Exists(Application.StartupPath+"\\BILL.DBF"))
		{
			MessageBox.Show("KINDLY COPY AND PAST THIS PROGRAM TO YOUR MEDI FOLDER.\nUNABLE TO LOCATE MEDI APPLICATION HERE\nUNABLE TO GET THE FIRM NAME FOR UPDATE.");
			return;
		}
		
		string m = "YOU HAD CHOOSE TO UPDATE\nWITH PHARMACEUTICAL SOFTWARE!\n ARE YOU SURE TO CONTINUE?\nPRESS YES TO CONTINUE OR PRESS NO TO CANCEL. ";
		DialogResult result = MessageBox.Show(this,m,"",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
		if(result ==DialogResult.Yes)
		{
			copy(Application.StartupPath,"..\\UPDATEBOTMEDIBACK"+DateTime.Now.ToString("ddMMyyyyhhmmss"));
			senddata();
			File.Delete(Application.StartupPath+"\\TOOLS.EXE");
            zipfile = "OBJ12.ZIP";
            downupd("las12");
			unzip();
		}
	}
    
	void dot12(object sender,EventArgs e)
	{
		if(!File.Exists(Application.StartupPath+"\\MEDI.EXE")||!File.Exists(Application.StartupPath+"\\BILL.DBF"))
		{
			MessageBox.Show("KINDLY COPY AND PAST THIS PROGRAM TO YOUR MEDI FOLDER.\nUNABLE TO LOCATE MEDI APPLICATION HERE\nUNABLE TO GET THE FIRM NAME FOR UPDATE.");
			return;
		}
		
		string m = "YOU HAD CHOOSE TO UPDATE\nWITH PHARMACEUTICAL SOFTWARE!\nARE YOU SURE TO CONTINUE?\nPRESS YES TO CONTINUE OR PRESS NO TO CANCEL. ";
		DialogResult result = MessageBox.Show(this,m,"",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
		if(result ==DialogResult.Yes)
		{
			copy(Application.StartupPath,"..\\UPDATEBOTMEDIBACK"+DateTime.Now.ToString("ddMMyyyyhhmmss"));
			senddata();
			File.Delete(Application.StartupPath+"\\TOOLS.EXE");
            zipfile = "DOT12.ZIP";
            downupd("dot12");
			unzip();
		}
	}	
	void ret(object sender,EventArgs e)
	{
		if(!File.Exists(Application.StartupPath+"\\MEDI.EXE")||!File.Exists(Application.StartupPath+"\\BILL.DBF"))
		{
			MessageBox.Show("KINDLY COPY AND PAST THIS PROGRAM TO YOUR MEDI FOLDER.\nUNABLE TO LOCATE MEDI APPLICATION HERE\nUNABLE TO GET THE FIRM NAME FOR UPDATE.");
			return;
		}
		
		string m = "YOU HAD CHOOSE TO UPDATE\nWITH PHARMACEUTICAL SOFTWARE!\nARE YOU SURE TO CONTINUE?\nPRESS YES TO CONTINUE OR PRESS NO TO CANCEL.";
		DialogResult result = MessageBox.Show(this,m,"",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
		if(result ==DialogResult.Yes)
		{
			copy(Application.StartupPath,"..\\UPDATEBOTMEDIBACK"+DateTime.Now.ToString("ddMMyyyyhhmmss"));
			senddata();
			File.Delete(Application.StartupPath+"\\TOOLS.EXE");
            zipfile = "RETAIL.ZIP";
            downupd("");
			unzip();
		}
	}
	void cust(object sender,EventArgs e)
	{
		if(!File.Exists(Application.StartupPath+"\\MEDI.EXE")||!File.Exists(Application.StartupPath+"\\BILL.DBF"))
		{
			MessageBox.Show("KINDLY COPY AND PAST THIS PROGRAM TO YOUR MEDI FOLDER.\nUNABLE TO LOCATE MEDI APPLICATION HERE\nUNABLE TO GET THE FIRM NAME FOR UPDATE.");
			return;
		}
		
		string m = "YOU HAD CHOOSE TO UPDATE\nWITHOUT CHANGEING INVOICE FORMAT!\nARE YOU SURE TO CONTINUE?\nPRESS YES TO CONTINUE OR PRESS NO TO CANCEL.";
		DialogResult result = MessageBox.Show(this,m,"",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
		if(result ==DialogResult.Yes)
		{
			copy(Application.StartupPath,"..\\UPDATEBOTMEDIBACK"+DateTime.Now.ToString("ddMMyyyyhhmmss"));
			senddata();
			File.Delete(Application.StartupPath+"\\TOOLS.EXE");
			zipfile = uploadinv();//not required in pre-set updates
			downupd("");
			unzip();
		}
	}

	void unzip()
	{
			Process process;
			if(zipfile!="")
			{
				process = new Process();
                
				process.StartInfo.FileName = Application.StartupPath+"\\fbzip.exe";
				process.StartInfo.Arguments =" -e -p "+zipfile+" \""+Application.StartupPath+"\"";
				process.Start();
				process.WaitForExit();
			}
			
			process = new Process();
			process.StartInfo.FileName = Application.StartupPath+"\\tools.exe";
			process.Start();
			process.WaitForExit();

			MessageBox.Show("UPDATE SUCCESSFULL!!\nNOTE:-PLEASE START MEDI AND DO FILE CREATION IN ALL OPTION MODE.\nPRESS OK TO EXIT THE PROGRAM" );
			Close();
	}
	string uploadinv()
	{
		string[] arr = new string[10];
		try
		{
			WebClient cl;
			cl = new WebClient();
			cl.Headers.Add("Content-Type","binary/octet-stream");
			cl.DownloadProgressChanged += new DownloadProgressChangedEventHandler(progchange);
			cl.DownloadFileCompleted += new AsyncCompletedEventHandler(progcompleted);
			byte[] result=cl.UploadFile("http://103.50.163.80/test.php","POST", Application.StartupPath+"\\INV.PRG");
			string s = Encoding.UTF8 .GetString (result,0,result.Length );	
			arr = s.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
			MessageBox.Show(s);
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.ToString());
		}
		return arr[7];
	}
	void downupd(string foldername)
	{
		
		List<string> filelist = new List<string>();
		
		if(zipfile=="")
		{
			filelist.Add("CHALP.O");
			filelist.Add("DATADICT.DBF");
			filelist.Add("DW.EXE");
			filelist.Add("inv.o");
			filelist.Add("INV.PRG");
			filelist.Add("M.BAT");
			filelist.Add("MAIL.EXE");
			filelist.Add("MEDI.O");
			filelist.Add("PPP.O");
			filelist.Add("PRINBAR.O");
			filelist.Add("RRR.O");
			filelist.Add("SSS.O");
			filelist.Add("STATE.DBF");
			filelist.Add("TMEDI.O");
			filelist.Add("TOOLS.EXE");
			filelist.Add("TTT.O");
			filelist.Add("XXX.O");
			filelist.Add("medi.exe");
			filelist.Add("GOVT.XLS");
            filelist.Add("FbZip.bin");

            foreach (string fn in filelist)
			{
				pbar.Value = 0;
                currentfile = fn;
				string filnam="http://medionline.co.in/update/"+foldername+"/"+fn;
				string location =Application.StartupPath+"\\"+fn;
				lbl3.Text= location;
				download(new Uri(filnam),location);
			}
		}
		else
		{
            filelist.Add(zipfile);
            filelist.Add("FBZip.exe");
			filelist.Add("FBZip.bin");
            foreach (string fn in filelist)
			{
				pbar.Value=0;
                currentfile = fn;
                //string filnam="http://103.50.163.80/"+fn;
                string filnam="http://medionline.co.in/update/"+fn;
                string location = Application.StartupPath+"\\"+fn;
                lbl3.Text = location;
                download(new Uri(filnam),location);
			}
		}
	}
	void download (Uri todownload,string savelocation)
	{
		WebClient cl = new WebClient();
        Stopwatch st = new Stopwatch();
        st.Start();
		cl.Headers.Add("Content-Type","binary/octet-stream");
		cl.DownloadProgressChanged += new DownloadProgressChangedEventHandler(progchange);
		cl.DownloadFileCompleted += new AsyncCompletedEventHandler(progcompleted);
		cl.DownloadFileAsync(todownload,savelocation );
		//without this, async download will go crazy and wont download anything
		while (cl.IsBusy) {Application.DoEvents(); }
	}
	void senddata()//procedure to send firmname to server 
	{
		DateTime fdate = File.GetLastWriteTime(Application.StartupPath+"\\MEDI.EXE");
		OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBase III;Data Source="+Application.StartupPath);
        conn.Open();
		string sql= "select FIRMNAME,ADDRESS,ADDRESS1,PHONEO,PHONER,PHONEM,COMPANY4 from bill";
		OleDbCommand cmd = new OleDbCommand(sql, conn);
		OleDbDataAdapter adap = new OleDbDataAdapter(cmd);
		DataSet ds = new DataSet();
		adap.Fill(ds,"data");
		conn.Close();
		StringWriter wrt = new StringWriter();
		ds.WriteXml(wrt);
		string xml=wrt.ToString();
		xml=xml.Replace("&amp;","");
		string exedate = fdate.ToString();
		
		string req = "http://medionline.co.in/update/my.php?data="+xml+"&exedate="+exedate;
		HttpWebRequest request = (HttpWebRequest)WebRequest.Create(req);
    	byte[] bytes;
    	bytes = Encoding.ASCII.GetBytes(xml);
    	request.ContentType = "text/xml; encoding='utf-8'";
    	request.ContentLength = bytes.Length;
    	request.Method = "POST";
    	Stream requestStream = request.GetRequestStream();
    	requestStream.Write(bytes, 0, bytes.Length);
    	requestStream.Close();
    	HttpWebResponse response;
    	response = (HttpWebResponse)request.GetResponse();
    	if (response.StatusCode == HttpStatusCode.OK)
    	{
    	    Stream responseStream = response.GetResponseStream();
    	    string responseStr = new StreamReader(responseStream).ReadToEnd();
    	    MessageBox.Show(responseStr);
    	}
	}

	void progchange(object sender, DownloadProgressChangedEventArgs e)
	{
		BeginInvoke((MethodInvoker) delegate {
        double bytesIn = double.Parse(e.BytesReceived.ToString());
        double totalBytes = double.Parse(e.TotalBytesToReceive.ToString());
        double percentage = bytesIn / totalBytes * 100;
        lbl2.Text = "Downloaded " + e.BytesReceived/1000 + "kb of " + e.TotalBytesToReceive/1000 +"kb";
        pbar.Value = int.Parse(Math.Truncate(percentage).ToString());
    	});
	}
	void progcompleted(object sender, AsyncCompletedEventArgs e)
	{
        BeginInvoke((MethodInvoker) delegate {
			pbar.Value=pbar.Maximum;
    	});
	}
	
	void copy( string sourcedir,string targetdir)
	{
		DirectoryInfo source = new DirectoryInfo(sourcedir);
		DirectoryInfo target = new DirectoryInfo(targetdir);
		copyall(source,target);		
	}
	void copyall(DirectoryInfo source,DirectoryInfo target)
	{
		Directory.CreateDirectory(target.FullName);

        // Copy each file into the new directory.
        foreach (FileInfo fi in source.GetFiles())
        {
            lbl2.Text="Copying "+target.FullName+"\\"+fi.Name;
			lbl2.Refresh();
            fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
        }
		
		// Copy each subdirectory using recursion.
        foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
        {
            DirectoryInfo nextTargetSubDir =
                target.CreateSubdirectory(diSourceSubDir.Name);
            copyall(diSourceSubDir, nextTargetSubDir);
        }
	}

}