using Laser_GPS_Converter_2.Config;
using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;
using System.Xml;
using System.Yaml;
using System.Yaml.Serialization;


namespace Laser_GPS_Converter_2
{

    public partial class Form1 : MaterialForm
	{
		DataSet tracks;

		public Form1()
		{
			InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Blue500, Primary.Blue900, Primary.Blue700, Accent.LightBlue200, TextShade.WHITE);
            ConfigYaml c = ConfigYaml.Instance;
            if(System.IO.File.Exists(c.Properties.DbPath))
            {
                Load_DB(ConfigYaml.Instance.Properties.DbPath);
            }
        }

		private void Form_Load(object sender, EventArgs e)
		{
			SetOffsetDefault();
		}


		private void SetOffsetDefault()
		{
	        //The times in the DB are stored locally, and GPX needs UTC
            //Set the 'default' timezone to that used by the computer since
            //it's probably not far from where the gps tracks were recorded
            TimeZone zone = TimeZone.CurrentTimeZone;
			n_Offset.Value = zone.GetUtcOffset(DateTime.Now).Hours;
		}

		private void btn_Load_Click(object sender, EventArgs e)
		{
            openFileDialog1.InitialDirectory = ConfigYaml.Instance.Properties.DbPath;
			list_Tracks.Items.Clear();
			DialogResult result = openFileDialog1.ShowDialog();

			if (result.Equals(DialogResult.Cancel))
			{
				return;
			}

            ConfigYaml.Instance.Properties.DbPath = openFileDialog1.FileName;
            ConfigYaml.Instance.Update();

            Load_DB(ConfigYaml.Instance.Properties.DbPath);
		}


        private void Load_DB(string path)
        {
            tracks = GetTracks(path);

            if (tracks == null)
            {
                //exception occurred
                MessageBox.Show("Error loading tracks!");
                return;
            }

            DataRowCollection dra = tracks.Tables["TrackPoint1"].Rows;
            for (int i = dra.Count - 1; i > -1; i--)
            {
                //Pretty print some details for each track to make them more easily identifiable
                //100000 factor worked out from checking the length of a known gpx record
                string l = dra[i][2] + ": " + dra[i][3].ToString().TrimEnd(' ') + " (" + Math.Round(Convert.ToDouble(dra[i][13]) / 100000, 2) + " km)";
                list_Tracks.Items.Add(l);
            }
        }

        //Gets the list of tracks - not the points from them, just which are available
        //Basically copying an MSDN example. I understand it, and it works, but there's probably a much shorter way of writing it.
		private DataSet GetTracks(string path)
		{
			OleDbConnection conn = null;
			DataSet t = new DataSet();
            string strAccessConn = @"Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + "; Jet OLEDB:Database Password=danger";
			string strAccessSelect = "SELECT * FROM TrackPoint1";
			try
			{
				conn = new OleDbConnection(strAccessConn);
			}
			catch (Exception ex)
			{
				Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
				return null;
			}

			try
			{
				OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, conn);
				OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
				conn.Open();
				myDataAdapter.Fill(t, "TrackPoint1");
			}
			catch (Exception ex)
			{
				Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
				return null;
			}
			finally
			{
				conn.Close();
			}
			return t;
		}

        //Finds the track to be exported in a horribly inefficient way.
        //Rewritten at some point to allow for multiple tracks to be exported at once;
        //I don't think I ever really tested that though.
		private void btn_Export_Click(object sender, EventArgs e)
		{
			if (list_Tracks.SelectedIndices.Count == 0)
			{
				MessageBox.Show("No tracks are selected.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			else
			{
                saveFileDialog1.InitialDirectory = ConfigYaml.Instance.Properties.ExportPath;
				DialogResult result = saveFileDialog1.ShowDialog();

				if (result == DialogResult.OK)
				{
                    ConfigYaml.Instance.Properties.ExportPath = System.IO.Path.GetDirectoryName(saveFileDialog1.FileName);
                    ConfigYaml.Instance.Update();

					int[] cNumbers = new int[list_Tracks.SelectedIndices.Count];
					for (int i = 0; i < list_Tracks.SelectedItems.Count; i++)
					{
						cNumbers[i] = Int32.Parse(list_Tracks.SelectedItems[i].ToString().Substring(0, list_Tracks.SelectedItems[i].ToString().IndexOf(':')));
					}
					ExportTracks(cNumbers, n_Offset.Value);
				}
			}
		}

        private void btn_ExportAll_Click(object sender, EventArgs e)
        {
            if( list_Tracks.Items.Count == 0)
            {
                MessageBox.Show("No database selected.", "Export all...", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                int[] cNumbers = new int[list_Tracks.Items.Count];
                for (int i = 0; i < list_Tracks.Items.Count; i++)
                {
                    cNumbers[i] = Int32.Parse(list_Tracks.Items[i].ToString().Substring(0, list_Tracks.Items[i].ToString().IndexOf(':')));
                }
                ExportAllTracks(cNumbers, n_Offset.Value);
            }
        }

        //Outputs the list of coordinates and other necessary data to a GPX file
        //Uses XMLWriter instead of the first version which was entirely appending strings manually
        //Probably not very efficient either, but it's still fast enough
		private void ExportTracks(int[] t, decimal offset)
		{
			XmlWriterSettings xs = new XmlWriterSettings();
			xs.Indent = true;
			xs.NamespaceHandling = NamespaceHandling.OmitDuplicates;

			XmlWriter writer = XmlWriter.Create(saveFileDialog1.FileName, xs);
			
			writer.WriteStartDocument();
            WriteGPXHeader(writer);

			foreach (int i in t)
			{
				DataSet trackPoints = GetTrackPoints(i);
				DataRowCollection dra = trackPoints.Tables["TrackPoint"].Rows;

                WriteGPX(writer, dra, offset);
			}

			writer.WriteEndDocument();
			writer.Close();
		}

        private void ExportAllTracks(int[] t, decimal offset)
        {
            folderBrowserDialog1.SelectedPath = ConfigYaml.Instance.Properties.ExportPath;
            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result != DialogResult.OK)
            {
                return;
            }

            ConfigYaml.Instance.Properties.ExportPath = folderBrowserDialog1.SelectedPath;
            ConfigYaml.Instance.Update();

            DataRowCollection draG = tracks.Tables["TrackPoint1"].Rows;
            int k = -1;
            foreach (int i in t)
            {
                k++;
                //Pretty print some details for each track to make them more easily identifiable
                //100000 factor worked out from checking the length of a known gpx record
                string fileName = draG[k][3].ToString().TrimEnd(' ').Replace(":", "-");

                XmlWriterSettings xs = new XmlWriterSettings();
                xs.Indent = true;
                xs.NamespaceHandling = NamespaceHandling.OmitDuplicates;

                XmlWriter writer = XmlWriter.Create(folderBrowserDialog1.SelectedPath + "\\" + fileName + ".gpx", xs);



                writer.WriteStartDocument();
                WriteGPXHeader(writer);

                DataSet trackPoints = GetTrackPoints(i);
                DataRowCollection dra = trackPoints.Tables["TrackPoint"].Rows;

                WriteGPX(writer, dra, offset);

                writer.WriteEndDocument();
                writer.Close();
            }
        }

        private void WriteGPXHeader(XmlWriter writer)
        {
            writer.WriteStartElement("gpx", @"http://www.topografix.com/GPX/1/1");
            writer.WriteAttributeString("creator", "Laser GPS Converter");
            //writer.WriteAttributeString("xmlns", @"http://www.topografix.com/GPX/1/0");
            writer.WriteAttributeString("version", "1.0");
            //writer.WriteAttributeString("xlmns", "xsi", null, @"http://www.w3.org/2001/XMLSchema-instance");
            writer.WriteAttributeString("xsi", "schemaLocation", @"http://www.w3.org/2001/XMLSchema-instance", @"http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd http://www.garmin.com/xmlschemas/GpxExtensions/v3 http://www.garmin.com/xmlschemas/GpxExtensionsv3.xsd http://www.garmin.com/xmlschemas/TrackPointExtension/v1 http://www.garmin.com/xmlschemas/TrackPointExtensionv1.xsd");
            writer.WriteAttributeString("xmlns", "gpxtpx", null, @"http://www.garmin.com/xmlschemas/TrackPointExtension/v1");
            writer.WriteAttributeString("xmlns", "gpxx", null, @"http://www.garmin.com/xmlschemas/GpxExtensions/v3");
        }

        private void WriteGPX(XmlWriter writer, DataRowCollection dra, decimal offset)
        {
            writer.WriteStartElement("trk");
            writer.WriteStartElement("trkseg");

            foreach (DataRow dr in dra)
            {
                writer.WriteStartElement("trkpt");
                //latitude
                if (dr[5].ToString().Trim().Equals("N"))
                    writer.WriteAttributeString("lat", dr[4].ToString().Replace(',', '.'));
                else
                    writer.WriteAttributeString("lat", "-" + dr[4].ToString().Replace(',', '.'));

                //longitude
                if (dr[3].ToString().Trim().Equals("E"))
                    writer.WriteAttributeString("lon", dr[2].ToString().Replace(',', '.'));
                else
                    writer.WriteAttributeString("lon", "-" + dr[2].ToString().Replace(',', '.'));

                //time
                string[] timebits = new string[3];
                timebits = dr[1].ToString().Trim().Split(':');
                for (int j = 0; j < 3; j++)
                {
                    timebits[j] = Int32.Parse(timebits[j]).ToString("D2");
                }
                string dt = dr[0].ToString().Substring(0, 10) + 'T' + String.Join(":", timebits);
                dt += offset > 0 ? '+' : '-';
                dt += ((int)offset).ToString("D2") + ':';
                dt += (offset - (int)offset == 0 ? "00" : "30");
                DateTime d = new DateTime(0, DateTimeKind.Local);
                d = DateTime.Parse(dt);

                writer.WriteElementString("time", d.ToUniversalTime().ToString("s") + 'Z');

                // Elevation
                if (Convert.ToDouble(dr[6].ToString()) / 100 > -200.00)
                {
                    writer.WriteElementString("ele", (Convert.ToDouble(dr[6].ToString()) / 100).ToString().Replace(',', '.'));
                }

                // Hearth Rate
                double hr = Convert.ToDouble(dr[7].ToString().Trim());
                if (hr > 0)
                {
                    writer.WriteStartElement("extensions");
                    writer.WriteStartElement("gpxtpx", "TrackPointExtension", null);
                    writer.WriteElementString("gpxtpx", "hr", null, hr.ToString());
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        //More 'borrowed' boilerplate code to access databases
        //Note that the password is apparently 'danger';
        //I don't think this is a security risk, all things considered.
        //Password retrieved via http://www.nirsoft.net/utils/accesspv.html
		private DataSet GetTrackPoints(int cNumber)
		{
			string strAccessConn = @"Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + ConfigYaml.Instance.Properties.DbPath + "; Jet OLEDB:Database Password=danger";
            //Add the appropriate columns here to be able to access them when exporting
            string strAccessSelect = "SELECT TrackPoint1.Track_Date, TrackPoint.TrackTime, TrackPoint.LongitudeN, TrackPoint.LonSign, TrackPoint.LatitudeN, TrackPoint.LatSign, TrackPoint.Alti, TrackPoint.HeartRate FROM TrackPoint, TrackPoint1 WHERE (((TrackPoint.cNumber)=[TrackPoint1].[cNumber]) AND ((TrackPoint1.cNumber)=" + cNumber + ")) ORDER BY TrackPoint.SerNO;";

			DataSet t = new DataSet();
			OleDbConnection myAccessConn = null;
			try
			{
				myAccessConn = new OleDbConnection(strAccessConn);
			}
			catch (Exception ex)
			{
				Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
			}

			try
			{

				OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
				OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

				myAccessConn.Open();
				myDataAdapter.Fill(t, "TrackPoint");

			}
			catch (Exception ex)
			{
				Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
			}
			finally
			{
				myAccessConn.Close();
			}

			return t;
		}

		private void list_Tracks_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_Tracks.SelectedItems.Count > 0)
			    UpdateTrackDetails();
		}

        //Shows more specific details about a selected track
		private void UpdateTrackDetails()
		{
			//Should improve this to show different details when multiple tracks selected - count, date range, total duration
			int i = list_Tracks.Items.Count - (list_Tracks.SelectedIndex + 1);
            Console.WriteLine(i);

			txt_Details.Clear();
			DataRow dr = tracks.Tables[0].Rows[i];
			txt_Details.AppendText("Started: " + dr[3].ToString().Trim());
			txt_Details.AppendText(Environment.NewLine + "Ended:   " + dr[5].ToString().Trim());
			txt_Details.AppendText(Environment.NewLine);

			DateTime t1 = ParseDate(dr, 3);
			DateTime t2 = ParseDate(dr, 5);
			TimeSpan duration = t2 - t1;
			txt_Details.AppendText(Environment.NewLine + "Duration: " + duration.ToString());

			txt_Details.AppendText(Environment.NewLine + "Distance: " + (Convert.ToDouble(dr[13]) / 100).ToString("n0") + " m");
		}

        //Converts the DB-stored datetime to a usable one
		private DateTime ParseDate(DataRow dr, int i)
		{
			string s = dr[i].ToString();
			s += n_Offset.Value > 0 ? '+' : '-';
			s += ((int)n_Offset.Value).ToString("D2") + ':';
			s += (n_Offset.Value - (int)n_Offset.Value == 0 ? "00" : "30");
			DateTime t = new DateTime(0, DateTimeKind.Local);
			t = DateTime.Parse(s);
			return t;
		}
	}
}
