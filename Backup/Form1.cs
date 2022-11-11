using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
//using Microsoft.ApplicationBlocks.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Design;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using System.Windows.Forms.PropertyGridInternal;
using System.Windows.Forms.ComponentModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DevExpress.Utils;
using DevExpress.Utils.Controls;
using DevExpress.Utils.Design;
using DevExpress.Utils.Drawing;
using DevExpress.Utils.Paint;
using DevExpress.Utils.Editors;
using DevExpress.Utils.WXPaint;
using DevExpress.Utils.Frames;
using DevExpress.XtraExport;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Persistent;
using DevExpress.XtraEditors.Container;
using DevExpress.XtraEditors.Design;
using DevExpress.XtraEditors.Drawing;
using DevExpress.XtraEditors.ListControls;
using DevExpress.XtraEditors.Mask;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraEditors.Registrator;
using DevExpress.XtraEditors.ViewInfo;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraVerticalGrid.Events;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Export;
using DevExpress.XtraGrid.Drawing;
//using DevExpress.ExpressApp.ConditionalAppearance;

//using Microsoft.Office.Interop.Excel;

//using AcAp = Autodesk.AutoCAD.Application;

namespace Employee
{
	

	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
//		Microsoft.Office.Interop.Excel.ApplicationClass ExcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
		DataSet ds = new DataSet();
		string FirstName = "First Name";
		string FatherName = "Father Name";
		string LastName = "Last Name";
		string FirstNameArabic = "First Name Arabic";
		string FatherNameArabic = "Father Name Arabic";
		string LastNameArabic = "Last Name Arabic";
		string EmpCode = "Employee Code";
		string Tax = "Tax#";
		string SocSec = "Soc Sec #";
		string EmpType = "Employment Type";
		string DepCode = "Department Code";
		string DepDesc = "Department Description";
		string PosCode = "Position Code";
		string PosDesc = "Position Description";
		string DOB = "DOB";
		string MaritalStatus = "Marital Status";
		string Gender = "Gender";
		string SpouseWork = "Spouse Work";
		string NatCode = "Nationality Code";
		string NatDesc = "Nationality Description";
		string Reg = "Registration #";
		string TimeAtt = "Time Attendance Badge#";
		string HiringDate = "Hiring Date";
		string TaxSince = "Tax Since";
		string NSSFSince = "NSSF Since";

		string Child1Name = "Child1 Name";
		string Child1Gender = "Child1 Gender";
		string Child1DOB = "Child1 DOB";
		string Child1TillDate = "Child1 on charge till date";
		string Child2Name = "Child2 Name";
		string Child2Gender = "Child2 Gender";
		string Child2DOB = "Child2 DOB";
		string Child2TillDate = "Child2 on charge till date";
		string Child3Name = "Child3 Name";
		string Child3Gender = "Child3 Gender";
		string Child3DOB = "Child3 DOB";
		string Child3TillDate = "Child3 on charge till date";
		string Child4Name = "Child4 Name";
		string Child4Gender = "Child4 Gender";
		string Child4DOB = "Child4 DOB";
		string Child4TillDate = "Child4 on charge till date";
		string Child5Name = "Child5 Name";
		string Child5Gender = "Child5 Gender";
		string Child5DOB = "Child5 DOB";
		string Child5TillDate = "Child5 on charge till date";

		string BasicSalaryUnit = "Basic Salary Unit";
		string PayFrequency = "Pay Frequency";
		string SalaryValue = "Salary Value";
		string CyOfSalary = "Cy of Salary";
		string TranspCode = "Transportation Code";
		string TranspDesc = "Transportation Desc";
		string TransValue = "Transportation Value";
		string TransUnit = "Transportation Unit D/P";
		string PaymentMethod = "Payment method";
		string BankAccount = "Bank Account";
		string Bank = "Bank";
		/// <summary>
		/// ///////////////////////////
		/// </summary>
		/// 
		string MothName = "Mother Name";
		string DisDate = "Termination Date";
		string Notes = "Notes";
		string Blood = "Blood Type";
		string ArMohafaza = "Mohafaza Arabic";
		string ArKadaa = "Kaza Arabic";
		string ArRegionTown = "City Arabic";
		string ArNeighborhood = "Neighborhood Arabic";
		string ArStreet = "Street Arabic";
		string ArBuilding = "Building Arabic";
		string ArFloor = "Floor Arabic";
		string ArPhone1 = "Phone1 Arabic";
		string NationalityDescArabic = "Nationality Desc Arabic";
		string Custom1Code = "Custom1 Code";
		string Custom1Desc = "Custom1 Description";
		string Custom2Code = "Custom2 Code";
		string Custom2Desc = "Custom2 Description";
		string PosDescArabic = "Position Desc Arabic";


		string MaritalStatusDate = "Marital Status Date";
		string HeafOfFamily = "Head Of Family";
		string NoChildrenAlloc = "No Children Allocation";
		string BankCode = "Bank Code";
		string BankDesc = "Bank";
		
		string GradeCode = "Grade Code";
		string GradeDesc = "Grade Description";
		string BranchCode = "Branch Code";
		string BranchDesc = "Branch Description";
		string CustomCr3 = "CustomCr3";
		string CustomCr4 = "CustomCr4";
		
		string PlaceOfBirth = "Place of Birth";
		string Nationality2Code = "Nationality2 Code";
		string Nationality2Desc = "Nationality2 Desc";
		string PassportNb = "Passport Number";
		string highCostOfLiving = "High cost of living";
		string IBAN = "IBAN";

		string Address1 = "Address1";
		string Phone = "Phone";
		string Mobile = "Mobile";
		string Email = "Email";
		string Address2 = "Secondary Address";
		string SecondaryPhone = "Secondary Phone";
		string SecondaryMobile = "Secondary Mobile";
		string SecondaryEmail = "Secondary Email";

		string IsStudent = "Is Student";
		string IsSmoker = "Is Smoker";

		string ContactReference = "Contact/Reference";
		string ContactName = "Contact Name";
		string ContactRelation = "Contact Relation";
		string ContactTel = "Contact Tel.";
		string ContactMobile = "Contact Mobile";

		string ArMotherName = "Mother Name (Arabic)";
		string IdCardNb = "ID Card Number";
		string ArKazaCard = "ID Card Kaza (Arabic)";
		string ArMohafazaCard = "ID Card Mohafaza (Arabic)";
		string ArRegisterPlace = "Register Place (Arabic)";
		string WorkPermitNb = "Work Permit Number";
		string WorkPermitDate = "Work Permit Date";
		string otherEmpNb = "other employer number";
		string otherEmpName = "other employer Name";
		string prevEmpName = "Previous Employer Name";
		string prevEmpNb = "Previous Employer Number";
		string prevEmpAddress = "Previous Employer Address";

		string ArEstateRegion = "Real Estate Region (Arabic)";
		string ArEstateNum = "Real Estate Number";
		string ArPhone2 = "Phone2"; //"Phone2 Arabic";
		string poBoxNb = "PO Box Number";
		string ArPoBoxRegion = "PO Box Region(Arabic)";
		string Fax = "Fax";

		string SpFirstName = "Spouse Name(Arabic)"; //"Spouse First Name";
		string SpLastName = "Spouse Maiden Name(Arabic)"; //"Spouse Last Name";
		string SpFatherName = "Spouse Father Name(Arabic)"; //"Spouse Father Name";
		string SpMotherName = "Spouse Mother Full Name (Maiden) (Arabic)";
		string SpNationality = "Spouse Nationality(Arabic)"; //"Spouse Nationality";
		string SpPOB = "Spouse Place of Birth(Arabic)";
		string SpDOB = "Spouse DOB";
		string SpIdCardNb = "Spouse ID Card Number";
		string SpRegisterNb = "Spouse Register Number";
		string SpRegisterPlace = "Spouse Register Place(Arabic)";
		string SpTax = "Spouse Tax Number";
		string SpCompanyName = "Spouse Company Name(Arabic)";
		string SpCompanyTax = "Spouse Company Tax Number";
		string SpWorkType = "Spouse Work Type(Arabic)";
		string publicOfficeName = "Public Office Name(Arabic)";
		string Spouse = "Spouse Name"; //spouse name
		string nssfExtDate = "NSSF Extension Date";
		string holdFamilyAlloc = "Hold Family Allocation";

		string IncomePayItemValue1 = "Income PayItem Value1";
		string IncomePayItemValue2 = "Income PayItem Value2";
		string DeductPayItemValue1 = "Deduct PayItem Value1";
		string DeductPayItemValue2 = "Deduct PayItem Value2";
		/// <summary>
		/// /////
		/// </summary>
		private DevExpress.XtraEditors.SimpleButton simpleButton1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
		public DevExpress.XtraGrid.GridControl dataGrid1;
		public DevExpress.XtraGrid.Views.Grid.GridView gridView3;
		private DevExpress.XtraEditors.SimpleButton simpleButton2;
		private DevExpress.XtraEditors.SimpleButton simpleButton3;
		private System.Data.SqlClient.SqlDataAdapter sqlDataAdapter1;
		private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
		private System.Data.SqlClient.SqlConnection sqlConnection1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Panel panel1;
		private DevExpress.XtraEditors.SimpleButton simpleButton4;
		private DevExpress.XtraEditors.SimpleButton simpleButton5;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public Form1()
		{
			
			//			LookUpEdit LookUp;
			//			LookUp = new LookUpEdit();
		
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton2 = new DevExpress.XtraEditors.SimpleButton();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.dataGrid1 = new DevExpress.XtraGrid.GridControl();
			this.gridView3 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.simpleButton3 = new DevExpress.XtraEditors.SimpleButton();
			this.sqlDataAdapter1 = new System.Data.SqlClient.SqlDataAdapter();
			this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
			this.sqlConnection1 = new System.Data.SqlClient.SqlConnection();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.panel1 = new System.Windows.Forms.Panel();
			this.simpleButton5 = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton4 = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView3)).BeginInit();
			this.panel1.SuspendLayout();
			this.SuspendLayout();
			// 
			// simpleButton1
			// 
			this.simpleButton1.Location = new System.Drawing.Point(384, 8);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(120, 23);
			this.simpleButton1.TabIndex = 1;
			this.simpleButton1.Text = "Import From Excel";
			this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// simpleButton2
			// 
			this.simpleButton2.Location = new System.Drawing.Point(16, 8);
			this.simpleButton2.Name = "simpleButton2";
			this.simpleButton2.TabIndex = 2;
			this.simpleButton2.Text = "Export";
			this.simpleButton2.Visible = false;
			this.simpleButton2.Click += new System.EventHandler(this.simpleButton2_Click);
			// 
			// gridView1
			// 
			this.gridView1.GridControl = null;
			this.gridView1.Name = "gridView1";
			// 
			// gridControl1
			// 
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(88, 88);
			this.gridControl1.MainView = this.gridView2;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.TabIndex = 0;
			// 
			// gridView2
			// 
			this.gridView2.GridControl = this.gridControl1;
			this.gridView2.Name = "gridView2";
			// 
			// dataGrid1
			// 
			this.dataGrid1.Dock = System.Windows.Forms.DockStyle.Fill;
			// 
			// dataGrid1.EmbeddedNavigator
			// 
			this.dataGrid1.EmbeddedNavigator.Name = "";
			this.dataGrid1.Location = new System.Drawing.Point(0, 0);
			this.dataGrid1.MainView = this.gridView3;
			this.dataGrid1.Name = "dataGrid1";
			this.dataGrid1.Size = new System.Drawing.Size(872, 306);
			this.dataGrid1.TabIndex = 3;
			this.dataGrid1.Text = "gridControl2";
			// 
			// gridView3
			// 
			this.gridView3.GridControl = this.dataGrid1;
			this.gridView3.GroupPanelText = "";
			this.gridView3.Name = "gridView3";
			this.gridView3.OptionsCustomization.AllowRowSizing = true;
			this.gridView3.OptionsView.ColumnAutoWidth = false;
			this.gridView3.OptionsView.ShowGroupPanel = false;
			// 
			// simpleButton3
			// 
			this.simpleButton3.Location = new System.Drawing.Point(512, 8);
			this.simpleButton3.Name = "simpleButton3";
			this.simpleButton3.Size = new System.Drawing.Size(112, 23);
			this.simpleButton3.TabIndex = 4;
			this.simpleButton3.Text = "Update Employees";
			this.simpleButton3.Click += new System.EventHandler(this.simpleButton3_Click);
			// 
			// sqlDataAdapter1
			// 
			this.sqlDataAdapter1.SelectCommand = this.sqlSelectCommand1;
			// 
			// sqlSelectCommand1
			// 
			this.sqlSelectCommand1.CommandText = "select *  from Emp";
			this.sqlSelectCommand1.Connection = this.sqlConnection1;
			// 
			// sqlConnection1
			// 
			this.sqlConnection1.ConnectionString = "workstation id=\"PC-PC\";packet size=4096;user id=sa;data source=\"pc-pc\";persist se" +
				"curity info=False;initial catalog=CleanPayrollTest2";
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.Filter = "XLS Files (*.xls)|*.xls|XLSX Files (*.xlsx)|*.xlsx";
			this.openFileDialog1.FilterIndex = 2;
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.simpleButton5);
			this.panel1.Controls.Add(this.simpleButton4);
			this.panel1.Controls.Add(this.simpleButton1);
			this.panel1.Controls.Add(this.simpleButton3);
			this.panel1.Controls.Add(this.simpleButton2);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.panel1.Location = new System.Drawing.Point(0, 306);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(872, 40);
			this.panel1.TabIndex = 5;
			// 
			// simpleButton5
			// 
			this.simpleButton5.Location = new System.Drawing.Point(760, 8);
			this.simpleButton5.Name = "simpleButton5";
			this.simpleButton5.Size = new System.Drawing.Size(104, 23);
			this.simpleButton5.TabIndex = 6;
			this.simpleButton5.Text = "Update Absence";
			this.simpleButton5.Click += new System.EventHandler(this.simpleButton5_Click);
			// 
			// simpleButton4
			// 
			this.simpleButton4.Location = new System.Drawing.Point(632, 8);
			this.simpleButton4.Name = "simpleButton4";
			this.simpleButton4.Size = new System.Drawing.Size(120, 23);
			this.simpleButton4.TabIndex = 5;
			this.simpleButton4.Text = "Update Transactions";
			this.simpleButton4.Click += new System.EventHandler(this.simpleButton4_Click);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(872, 346);
			this.Controls.Add(this.dataGrid1);
			this.Controls.Add(this.panel1);
			this.Name = "Form1";
			this.Text = "Employee List";
			this.Load += new System.EventHandler(this.Form1_Load);
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView3)).EndInit();
			this.panel1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());

		}
		string ServerName;
		string Database;
		string UserID;
		string Password;

		private void Form1_Load(object sender, System.EventArgs e)
		{
			Form2 frm = new Form2();
			frm.ShowDialog();
			if (!frm.OK)
			{
				Application.Exit();
			}
			ServerName = frm.textBox1.Text.Trim();
			Database = frm.textBox2.Text.Trim();
			UserID = frm.textBox3.Text.Trim();
			Password = frm.textBox4.Text.Trim();
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			try
			{
				if(this.openFileDialog1.ShowDialog() == DialogResult.OK)
				{
					OleDbConnection con = new OleDbConnection();
					con.ConnectionString = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + this.openFileDialog1.FileName + ";Extended Properties=\"Excel 12.0;HDR=Yes\"";
					con.Open();
					DataTable dtSchema;
					dtSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
					OleDbCommand Command = new OleDbCommand ("select * FROM [" + dtSchema.Rows[0]["TABLE_NAME"].ToString() + "]", con);
					OleDbDataAdapter da = new OleDbDataAdapter(Command);
					DataSet ds = new DataSet ();
					da.Fill(ds);
					dataGrid1.DataSource = ds.Tables[0];
					con.Close();
					//--- removing empty lines
					for (int i = this.gridView3.RowCount - 1; i >= 0; i--)
					{
						try
						{
							if (this.gridView3.GetRowCellValue(i, this.gridView3.Columns[FirstName]) == DBNull.Value || this.gridView3.GetRowCellValue(i, this.gridView3.Columns[FirstName]).ToString().Trim().Length == 0)
							{
								this.gridView3.DeleteRow(i);
							}
						}
						catch(Exception ex)
						{
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
		
		
		private void simpleButton2_Click(object sender, System.EventArgs e)
		{
			IExportProvider provider = new ExportXlsProvider("C:\\Users\\pc\\Documents\\Emp1.xls");
			BaseExportLink link = this.gridView3.CreateExportLink(provider);
			link.ExportTo(true);
		}


		private void simpleButton3_Click(object sender, System.EventArgs e)
		{	
			//=====================
			//===========check CODE
			//=====================
			for (int i = 0; i < gridView3.RowCount - 1; i++)
			{
				if (gridView3.GetRowCellValue(i, gridView3.Columns[EmpCode]).ToString() == "")
				{
					MessageBox.Show("Employee Code line " + i + " can't be empty", MessageBoxButtons.OK.ToString());
					return;
				}
			}

			for (int i = 0; i < gridView3.RowCount - 1; i++)
			{
				for (int j = i + 1; j < gridView3.RowCount; j++)
				{
					if (gridView3.GetRowCellValue(i, gridView3.Columns[EmpCode]).ToString() == gridView3.GetRowCellValue(j, gridView3.Columns[EmpCode]).ToString())
					{
						MessageBox.Show("Duplicate Employee Code line " + i + " and " + j, MessageBoxButtons.OK.ToString());
						return;
					}
				}
			}
			//=====================

			//=====================
			//=========check Badge#
			//=====================
			for (int i = 0; i < gridView3.RowCount - 1; i++)
			{
				if (gridView3.GetRowCellValue(i, gridView3.Columns[TimeAtt]).ToString() == "")
				{
					MessageBox.Show("Employee Badge# line " + i+ " can't be empty", MessageBoxButtons.OK.ToString());
					return;
				}
			}

			for (int i = 0; i < gridView3.RowCount - 1; i++)
			{
				for (int j = i + 1; j < gridView3.RowCount; j++)
				{
					if (gridView3.GetRowCellValue(i, gridView3.Columns[TimeAtt]).ToString() == gridView3.GetRowCellValue(j, gridView3.Columns[TimeAtt]).ToString())
					{
						MessageBox.Show("Duplicate Employee Code line " + i + " and " + j, MessageBoxButtons.OK.ToString());
						return;
					}
				}
			}
			//=====================

			for (int i = 0; i < gridView3.RowCount; i++)
			{
				if (gridView3.GetRowCellValue(i, gridView3.Columns[FirstName]) == DBNull.Value || gridView3.GetRowCellValue(i, gridView3.Columns[FirstName]).ToString() == "")
				{
					MessageBox.Show("First Name in the row number " + i + "  can't be empty", MessageBoxButtons.OK.ToString());
					return;
				}
	
				if (gridView3.GetRowCellValue(i, gridView3.Columns[LastName]).ToString() == "")
				{
					MessageBox.Show("Last Name in the row "+ i+" can't be empty", MessageBoxButtons.OK.ToString());
					return;
				}

//				object emp_code = gridView3.GetRowCellValue(i, gridView3.Columns["Employee Code"]);

				if (gridView3.GetRowCellValue(i, gridView3.Columns[EmpCode]).ToString() == "")
				{
					MessageBox.Show("Employee Code line " + i+ " can't be empty", MessageBoxButtons.OK.ToString());
					return;
				}

				string cellvalue = gridView3.GetRowCellValue(i, gridView3.Columns[EmpType]).ToString();
				if (cellvalue != "OWNER")
				{
					if (cellvalue != "PA")
					{
						if( cellvalue != "CO")
						{
							if( cellvalue != "PT")
							{
								if(cellvalue != "FULL TIME")
								{
									if (cellvalue != "owner")
									{
										if (cellvalue != "pa")
										{
											if( cellvalue != "co")
											{
												if( cellvalue != "pt")
												{
													if (cellvalue.ToUpper() != "FT")
													{
														if(cellvalue != "full time")
														{
															MessageBox.Show("Employment Type line " + i + " :" + "\n" + "'OW' : Owner" + "\n" + "'PA' : Partner" + "\n" + "'CO' : Contractual" + "\n" + "'PT' : Part Time" + "\n" + "'FT' : Full Time" + "\n");
															return;
														}
													}
												}
											}
										}
									}
								}
							}
						}
					}
				}

				if (gridView3.GetRowCellValue(i, gridView3.Columns[DepCode]).ToString() == "")
				{
					MessageBox.Show("'Department Code' line " + i+ " can't be empty", MessageBoxButtons.OK.ToString());
					return;
				}
				if (gridView3.GetRowCellValue(i, gridView3.Columns[DepDesc]).ToString() == "")
				{
					MessageBox.Show("'Department Description' line " + i+ " can't be empty", MessageBoxButtons.OK.ToString());
					return;
				}
				if (gridView3.GetRowCellValue(i, gridView3.Columns[PosDesc]).ToString() == "")
				{
					MessageBox.Show("'Position Description' line " + i+ " can't be empty", MessageBoxButtons.OK.ToString());
					return;
				}
				//				if (gridView3.GetRowCellValue(i, gridView3.Columns[PosDesc]).ToString() == "")
				//				{
				//					MessageBox.Show("'Position Description' line " + i+ " can't be empty", MessageBoxButtons.OK.ToString());
				//					return;
				//					Application.Run();
				//				}


//				if (gridView3.GetRowCellValue(i, gridView3.Columns[DOB]).ToString() == "") 
//				{
//					MessageBox.Show("Date Of Birth of the row number " + i + " can't be empty", MessageBoxButtons.OK.ToString());
//					return;
//				}

				cellvalue = (gridView3.GetRowCellValue(i, gridView3.Columns[MaritalStatus]).ToString());
						
				if (cellvalue != "S") 
				{
					if (cellvalue != "M")  
					{
						if (cellvalue != "D") 
						{
							if (cellvalue != "W")
							{
								if (cellvalue != "C")
								{
									if (cellvalue != "V")
									{
										if (cellvalue != "s")
										{
											if (cellvalue != "m")
											{
												if (cellvalue != "d")
												{
													if (cellvalue != "w")
													{
														if (cellvalue != "c")
														{
															if (cellvalue != "v")
															{
																MessageBox.Show("Marital Status line " + i + " : 'S' - 'M' - 'D' - 'W' - 'C'");
																return;
															}
														}
													}
												}
											}
										}
									}
								}
							}									
						}
					}
				}

				cellvalue = (gridView3.GetRowCellValue(i, gridView3.Columns[Gender]).ToString());
								
				if (cellvalue != "M")
				{
					if(cellvalue != "m")
					{
						if(cellvalue != "f")
						{
							if(cellvalue != "F")
							{
								MessageBox.Show("Gender line " + i + " must contain the values : 'F' or 'M' ");
								return;
							}
						}
					}
				}
			
				cellvalue = (gridView3.GetRowCellValue(i,gridView3.Columns[BasicSalaryUnit]).ToString());
			
				if (cellvalue != "M")  
				{
					if (cellvalue != "D") 
					{
						if (cellvalue != "H")
						{
							if (cellvalue != "m")
							{
								if (cellvalue != "d")
								{
									if (cellvalue != "h")
									{
										MessageBox.Show("Basic Salary Unit line " + i + " must contain the values : 'M' or 'D' or 'H'  ");
										return;
									}
								}
							}
						}								
					}
				}
			
				cellvalue = gridView3.GetRowCellValue(i,gridView3.Columns[PayFrequency]).ToString();
				if (cellvalue != "M") 
				{
					if (cellvalue != "W")  
					{
						if (cellvalue != "BW") 
						{
							if (cellvalue != "SM")
							{
								if (cellvalue != "m")
								{
									if (cellvalue != "w")
									{
										if (cellvalue != "bw")
										{
											if (cellvalue != "sm")
											{
												MessageBox.Show("Pay Frequency line " + i + " : 'M' - 'W' - 'BW' - 'SM' ");
												return;
											}
										}
									}
								}
							}									
						}
					}
				}
				/////////////////////////////////////// child 1

				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1Name]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1DOB]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child1 DOB' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1Gender]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child1 Gender' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child1 on charge till' line " + i );
//						return;
//					}
				}
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1DOB]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1Name]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child1 Name' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1Gender]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child1 Gender' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child1 on charge till' line " + i );
//						return;
//					}
				}
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1Gender]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1Name]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child1 Name' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1DOB]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child1 DOB' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child1 on charge till' line " + i );
//						return;
//					}
				}
//				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1TillDate]).ToString() != "")
//				{
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1Name]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child1 Name' line " + i );
//						return;
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1DOB]).ToString() == "") 
//					{
//						MessageBox.Show("You have to fill 'Child1 DOB' line " + i );
//						return;
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1Gender]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child1 Gender till' line " + i );
//						return;
//					}
//				}
				////////////////////////////////////////////////////////// child 2
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2Name]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2DOB]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child2 DOB' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2Gender]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child2 Gender' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child2 on charge till' line " + i );
//						return;
//					}
				}
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2DOB]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2Name]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child2 Name' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2Gender]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child2 Gender' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child2 on charge till' line " + i );
//						return;
//					}
				}
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2Gender]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2Name]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child2 Name' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2DOB]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child2 DOB' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child2 on charge till' line " + i );
//						return;
//					}
				}
//				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2TillDate]).ToString() != "")
//				{
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2Name]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child2 Name' line " + i );
//						return;
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2DOB]).ToString() == "") 
//					{
//						MessageBox.Show("You have to fill 'Child2 DOB' line " + i );
//						return;
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2Gender]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child2 Gender till' line " + i );
//						return;
//					}
//				}
				////////////////////////////////////////////////////////////// child 3
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3Name]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3DOB]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child3 DOB' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3Gender]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child3 Gender' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child3 on charge till' line " + i );
//						return;
//					}
				}
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3DOB]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3Name]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child3 Name' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3Gender]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child3 Gender' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child3 on charge till' line " + i );
//						return;
//					}
				}
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3Gender]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3Name]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child3 Name' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3DOB]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child3 DOB' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child3 on charge till' line " + i );
//						return;
//					}
				}
//				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3TillDate]).ToString() != "")
//				{
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3Name]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child3 Name' line " + i );
//						return;
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3DOB]).ToString() == "") 
//					{
//						MessageBox.Show("You have to fill 'Child3 DOB' line " + i );
//						return;
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3Gender]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child3 Gender till' line " + i );
//						return;
//					}
//				}
				//////////////////////////////////////////////child 4
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4Name]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4DOB]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child4 DOB' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4Gender]).ToString() == "") 
//					{
//						MessageBox.Show("You have to fill 'Child4 Gender' line " + i );
//						return;
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child4 on charge till' line " + i );
//						return;
//					}
				}
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4DOB]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4Name]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child4 Name' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4Gender]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child4 Gender' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child4 on charge till' line " + i );
//						return;
//					}
				}
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4Gender]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4Name]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child4 Name' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4DOB]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child4 DOB' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child4 on charge till' line " + i );
//						return;
//					}
				}
//				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4TillDate]).ToString() != "")
//				{
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4Name]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child4 Name' line " + i );
//						return;
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4DOB]).ToString() == "") 
//					{
//						MessageBox.Show("You have to fill 'Child4 DOB' line " + i );
//						return;
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4Gender]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child4 Gender till' line " + i );
//						return;
//					}
//				}
				//////////////////////////////////////////////child 5
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5Name]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5DOB]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child5 DOB' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5Gender]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child5 Gender' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child5 on charge till' line " + i );
//						return;
//					}
				}
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5DOB]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5Name]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child5 Name' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5Gender]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child5 Gender' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child5 on charge till' line " + i );
//						return;
//					}
				}
				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5Gender]).ToString() != "")
				{
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5Name]).ToString() == "")
					{
						MessageBox.Show("You have to fill 'Child5 Name' line " + i );
						return;
					}
					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5DOB]).ToString() == "") 
					{
						MessageBox.Show("You have to fill 'Child5 DOB' line " + i );
						return;
					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5TillDate]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child5 on charge till' line " + i );
//						return;
//					}
				}
//				if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5TillDate]).ToString() != "")
//				{
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5Name]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child5 Name' line " + i );
//						return;
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5DOB]).ToString() == "") 
//					{
//						MessageBox.Show("You have to fill 'Child5 DOB' line " + i );
//						return;
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5Gender]).ToString() == "")
//					{
//						MessageBox.Show("You have to fill 'Child5 Gender till' line " + i );
//						return;
//					}
//				}



//				if (TranspDesc.Length > 0)
//				{
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[TranspDesc]).ToString() != "")
//					{
//						if (gridView3.GetRowCellValue(i,gridView3.Columns[TransUnit]).ToString() == "")
//						{
//							MessageBox.Show("You have to fill 'Transportation Unit' line " + i );
//							return;
//						}
//					}
//					if (gridView3.GetRowCellValue(i,gridView3.Columns[TransUnit]).ToString() != "")
//					{
//						if (gridView3.GetRowCellValue(i,gridView3.Columns[TranspDesc]).ToString() == "")
//						{
//							MessageBox.Show("You have to fill 'Transportation Description' line " + i );
//							return;
//						}
//					}
//				}
			}

			string sql = "insert into Emp ([Name], [FName], [MName], [LName], [Code], [TaxNb], [SSN], [EmploymentType], [Department], [DOB], [MarStat], [Gender], [RegNum], [BadgeNum], [Inactive], [SpouseWork], [Decease], [Discharge], [ArEmpName], [ArFatherName], [ArLastName], [SpouseCharge], [NoFaBen], [Default], [HoldFamilyAlloc], [EOSWithdrawDt], [BasicPayUnit], [PayFrequency], [ValuePerUnit], [CurrID], [GenPolID], [TaxPolID], [PayPolicyID], [CurrRoundID], [UseClassDef], [HireDate], [TaxSince], [HireSince], [Spouse], [SpName], [SpLName], [SpFaName], [SpMoNameL], [SpNationality], [SpPOB], [SpDOB], [PositionA], [Nationality1], AccountNum, AccountCy, SAccountNum, SAccountCy, CustomCr1, CustomCr2, Address1, Address2, Phone, Mobile, DisDate, DisCode, MothName, Notes, Blood, ArMohafaza, ArKadaa, ArRegionTown, ArNeighborhood, ArStreet, ArEstateRegion, ArEstateNum, ArBuilding, ArFloor, ArPhone1, ArPhone2, TAPolID, Custom3, Custom4, POB, PassportNum, HighCostOfLiving, IBAN, email, isStudent, isSmoker, ArMothName, ArIDNum, ArSijilKazaa, ArSijilMohafaza, ArSijilPlace, SpIDNum, ArSijilNum, ArNationality, Branch) "
			+ "values (@Name, @firstname, @middlename, @lastname, @code, @TaxNb, @SSN, @EmpType, @Department, cast(@DOB as datetime), @MarStat, @Gender, @RegNum, @BadgeNum, @Inactive, @SpouseWork, @Decease, @Discharge, @ArEmpName, @ArFatherName, @ArLastName, @SpouseCharge, @NoFaBen, @Default, @HoldFamilyAlloc, cast(@EOSWithdrawDt as datetime), @BasicPayUnit, @PayFrequency, @ValuePerUnit, @CurrID, @GenPolID, @TaxPolID, @PayPolicyID, @CurrRoundID, @UseClassDef, cast(@HireDate as datetime), cast(@TaxSince as datetime), cast(@HireSince as datetime), @Spouse, @SpName, @SpLName, @SpFaName, @SpMoNameL, @SpNationality, @SpPOB, cast(@SpDOB as datetime), @PositionA, @Nationality1, @AccountNum, @CurrID, @AccountNum, @CurrID, @CustomCr1, @CustomCr2, @Address1, @Address2, @Phone, @Mobile, cast(@DisDate as datetime), @DisCode, @MothName, @Notes, @Blood, @ArMohafaza, @ArKadaa, @ArRegionTown, @ArNeighborhood, @ArStreet, @ArEstateRegion, @ArEstateNum, @ArBuilding, @ArFloor, @ArPhone1, @ArPhone2, @TAPolID, @CustomCr3, @CustomCr4, @POB, @PassportNum, @HighCostL, @IBAN, @email, @IsStudent, @IsSmoker, @ArMotherName, @ArIDNum, @ArSijilKazaa, @ArSijilMohafaza, @ArSijilPlace, @SpIDNum, @ArSijilNum, @ArNationalityRegisterP, @BranchID)";
	
			string sqlDep = @"Insert into tbDepartment ([Title], [Abbreviation]) values (@Desc, @DepCode)";
			string sqlBranch = @"Insert into tbBranch ([Title], [Abbreviation]) values (@BranchDesc, @BranchCode)";
			string sqlPos = @"Insert into tbPosition ([Title], [Abbreviation], ArabicTitle) values (@PosDesc, @PosCode, @ArabicTitle)";
			string sqlNat = @"Insert into tbNationality ([Title], [Abbreviation], [ArabicTitle]) values (@NatDesc, @NatCode, @ArabicTitle)";
			
			string sqlEmpFamily = @"Insert into EmpFamily ([EmpID], [Name], [DOB], [Gender], [OnCharge], [ChExemEndDate]) values (@EmpID, @Name, cast(@DOB as datetime), @Gender, @OnCharge, cast(@ChExemEndDate as datetime))";
			string sqlPayItems = @"Insert into PayItems ([ItmType],[ItmDesc],[ItmAbb],[ItmCalc],[Rate],[CurrID],[GlobalSalary],[Special],[Taxable],[CalcEOS],[CalcFA],[CalcSick],[CalcXtraEOS],[LeaveUpdate],[Reloan],[InEmpCost],[Printable],[DedTax],[Inactive],[YearlyPayVac],[Loan],[Provision],[R6],[Overtime],[BasicPay],[Istaxable],[HighCost],[IsTransport],[School],[ExtraMonth],[HighCostOfLiving]) values (@ItmType,@ItmDesc,@ItmAbb,@ItmCalc,@Rate,@CurrID,@GlobalSalary,@Special,@Taxable,@CalcEOS,@CalcFA,@CalcSick,@CalcXtraEOS,@LeaveUpdate,@Reloan,@InEmpCost,@Printable,@DedTax,@Inactive,@YearlyPayVac,@Loan,@Provision,@R6,@Overtime,@BasicPay,@Istaxable,@HighCost,@IsTransport,@School,@ExtraMonth,@HighCostOfLiving)";
			string bankAccount = "insert into tbbankaccounts (currid, bankid, bankaccount, accountdesc) values (@currid, @bankid, @bankaccount, @accountdesc)";
			string bank = "insert into tbbank (bank, applytransfer) values (@bank, @applytransfer)";
			string bankSymbols = "insert into tbbankcurrsymbols (bankid, ourcurrency, banksymbol) values (@bankid, @ourcurrency, @banksymbol)";
			string payPolicy = "insert into tbpaypolicy (title, regpaymethod, specialpaymethod, frombankidregular, frombankidspecial) values (@title, @regpaymethod, @specialpaymethod, @frombankidregular, @frombankidspecial)";
			string sqlCustom1 = "insert into tbCustomCr1 (Title, Abbreviation) values (@Title, @Abbreviation)";
			string sqlCustom2 = "insert into tbCustomCr2 (Title, Abbreviation) values (@Title, @Abbreviation)";

			string sqlMaritalStatus = "insert into emphistory (empid, effdate, marstat, spousework, headoffamily, nocfaben) values (@emID, @MaritalStatDate, @MaritalStatSymbol, @SpouseWork, @headoffamily, 0)";

			string sqlPayItem1 = "insert into EmpRegPayTypes (EmpID, PayTypeID, CurrID, Value, PaymentFreq, Type) values (@emID1, 69, 149, @piValue1, 'EP', 'Normal')";
			string sqlPayItem2 = "insert into EmpRegPayTypes (EmpID, PayTypeID, CurrID, Value, PaymentFreq, Type) values (@emID2, 53, 149, @piValue2, 'EP', 'Normal')";
			string sqlPayItemDed1 = "insert into EmpRegPayTypes (EmpID, PayTypeID, CurrID, Value, PaymentFreq, Type) values (@emIDd1, 70, 149, @pidValue1, 'EP', 'Normal')";
			string sqlPayItemDed2 = "insert into EmpRegPayTypes (EmpID, PayTypeID, CurrID, Value, PaymentFreq, Type) values (@emIDd2, 71, 149, @pidValue2, 'EP', 'Normal')";

			using(SqlConnection connection = new SqlConnection("user id=" + UserID + ";Password=" + Password +";data source=" + ServerName + ";persist security info=True;initial catalog=" + Database))
			{
				using(SqlCommand command = new SqlCommand(sql, connection))
				{	
					try
					{
						SqlCommand cmDep = new SqlCommand(sqlDep, connection);
						
						SqlCommand cmdEmpFamily = new SqlCommand(sqlEmpFamily, connection);

						SqlCommand cmdPayItems = new SqlCommand(sqlPayItems, connection);

						SqlCommand cmdEmpPayItems = new SqlCommand("insert into empregpaytypes (empid, paytypeid, currid, value, paymentfreq, type) values (@empid, @paytypeid, @currid, @value, @paymentfreq, @type)", connection);

						SqlCommand cmdMaritalStatus = new SqlCommand(sqlMaritalStatus, connection);
						cmdMaritalStatus.Parameters.Add("@emID", SqlDbType.Int);
						cmdMaritalStatus.Parameters.Add("@MaritalStatDate", SqlDbType.DateTime);
						cmdMaritalStatus.Parameters.Add("@SpouseWork", SqlDbType.Bit);
						cmdMaritalStatus.Parameters.Add("@MaritalStatSymbol", SqlDbType.NVarChar);
						cmdMaritalStatus.Parameters.Add("@headoffamily", SqlDbType.Bit);
						
						SqlCommand cmdPayItem1 = new SqlCommand(sqlPayItem1, connection);
						cmdPayItem1.Parameters.Add("@emID1", SqlDbType.Int);
						cmdPayItem1.Parameters.Add("@piValue1", SqlDbType.Decimal);

						SqlCommand cmdPayItem2 = new SqlCommand(sqlPayItem2, connection);
						cmdPayItem2.Parameters.Add("@emID2", SqlDbType.Int);
						cmdPayItem2.Parameters.Add("@piValue2", SqlDbType.Decimal);
						
						SqlCommand cmdPayItemDed1 = new SqlCommand(sqlPayItemDed1, connection);
						cmdPayItemDed1.Parameters.Add("@emIDd1", SqlDbType.Int);
						cmdPayItemDed1.Parameters.Add("@pidValue1", SqlDbType.Decimal);

						SqlCommand cmdPayItemDed2 = new SqlCommand(sqlPayItemDed2, connection);
						cmdPayItemDed2.Parameters.Add("@emIDd2", SqlDbType.Int);
						cmdPayItemDed2.Parameters.Add("@pidValue2", SqlDbType.Decimal);

						SqlCommand cmdBank = new SqlCommand(bank, connection);
						cmdBank.Parameters.Add("@bank", SqlDbType.NVarChar);
						cmdBank.Parameters.Add("@applytransfer", SqlDbType.Bit);

						SqlCommand cmdBankAccount = new SqlCommand(bankAccount, connection);
						cmdBankAccount.Parameters.Add("@currid", SqlDbType.SmallInt);
						cmdBankAccount.Parameters.Add("@bankid", SqlDbType.Int);
						cmdBankAccount.Parameters.Add("@bankaccount", SqlDbType.VarChar);
						cmdBankAccount.Parameters.Add("@accountdesc", SqlDbType.NVarChar);

						SqlCommand cmdBankSymbols = new SqlCommand(bankSymbols, connection);
						cmdBankSymbols.Parameters.Add("@bankid", SqlDbType.Int);
						cmdBankSymbols.Parameters.Add("@ourcurrency", SqlDbType.SmallInt);
						cmdBankSymbols.Parameters.Add("@banksymbol", SqlDbType.NVarChar);

						SqlCommand cmdPayPolicy = new SqlCommand(payPolicy, connection);
						cmdPayPolicy.Parameters.Add("@title", SqlDbType.NVarChar);
						cmdPayPolicy.Parameters.Add("@regpaymethod", SqlDbType.NVarChar);
						cmdPayPolicy.Parameters.Add("@specialpaymethod", SqlDbType.NVarChar);
						cmdPayPolicy.Parameters.Add("@frombankidregular", SqlDbType.Int);
						cmdPayPolicy.Parameters.Add("@frombankidspecial", SqlDbType.Int);

						cmdEmpPayItems.Parameters.Add("@empid", SqlDbType.Int);
						cmdEmpPayItems.Parameters.Add("@paytypeid", SqlDbType.Int);
						cmdEmpPayItems.Parameters.Add("@currid", SqlDbType.SmallInt);
						cmdEmpPayItems.Parameters.Add("@value", SqlDbType.Decimal);
						cmdEmpPayItems.Parameters.Add("@paymentfreq", SqlDbType.NVarChar);
						cmdEmpPayItems.Parameters.Add("@type", SqlDbType.NVarChar);

						SqlCommand cmdBranch = new SqlCommand(sqlBranch,connection);
						cmdBranch.Parameters.Add("@BranchDesc", SqlDbType.NVarChar);
						cmdBranch.Parameters.Add("@BranchCode", SqlDbType.NVarChar);

						SqlCommand cmdPos = new SqlCommand(sqlPos,connection);
						cmdPos.Parameters.Add("@PosDesc", SqlDbType.NVarChar);
						cmdPos.Parameters.Add("@PosCode", SqlDbType.NVarChar);
						cmdPos.Parameters.Add("@ArabicTitle", SqlDbType.NVarChar);

						SqlCommand cmdNat = new SqlCommand(sqlNat,connection);
						cmdNat.Parameters.Add("@NatDesc", SqlDbType.NVarChar);
						cmdNat.Parameters.Add("@NatCode", SqlDbType.NVarChar);
						cmdNat.Parameters.Add("@ArabicTitle", SqlDbType.NVarChar);

						SqlCommand cmdCustom1 = new SqlCommand(sqlCustom1, connection);
						cmdCustom1.Parameters.Add("@Title", SqlDbType.NVarChar);
						cmdCustom1.Parameters.Add("@Abbreviation", SqlDbType.NVarChar);

						SqlCommand cmdCustom2 = new SqlCommand(sqlCustom2, connection);
						cmdCustom2.Parameters.Add("@Title", SqlDbType.NVarChar);
						cmdCustom2.Parameters.Add("@Abbreviation", SqlDbType.NVarChar);

						command.Parameters.Add("@Name" , SqlDbType.NVarChar);
						command.Parameters.Add("@firstname" , SqlDbType.NVarChar);
						command.Parameters.Add("@middlename" , SqlDbType.NVarChar);
						command.Parameters.Add("@lastname" , SqlDbType.NVarChar);
						command.Parameters.Add("@code" , SqlDbType.NVarChar);
						command.Parameters.Add("@TaxNb" , SqlDbType.NVarChar);
						command.Parameters.Add("@SSN" , SqlDbType.NVarChar);
						command.Parameters.Add("@EmpType" , SqlDbType.Int);
						command.Parameters.Add("@Department", SqlDbType.Int);
						command.Parameters.Add("@PositionA", SqlDbType.Int);
						command.Parameters.Add("@Nationality1", SqlDbType.Int);
						
						cmDep.Parameters.Add("@Desc", SqlDbType.VarChar);
						cmDep.Parameters.Add("@DepCode", SqlDbType.VarChar);

						command.Parameters.Add("@DOB" , SqlDbType.NVarChar);
						command.Parameters.Add("@MarStat" , SqlDbType.VarChar);
						command.Parameters.Add("@Gender" , SqlDbType.VarChar);
						command.Parameters.Add("@RegNum" , SqlDbType.VarChar);
						command.Parameters.Add("@BadgeNum" , SqlDbType.VarChar);
						command.Parameters.Add("@GenPolID", SqlDbType.Int);
						command.Parameters.Add("@Inactive", SqlDbType.VarChar);	
						command.Parameters.Add("@SpouseWork", SqlDbType.Bit);	
						command.Parameters.Add("@Decease", SqlDbType.VarChar);	
						command.Parameters.Add("@Discharge", SqlDbType.VarChar);	
						command.Parameters.Add("@IsStudent", SqlDbType.VarChar);	
						command.Parameters.Add("@IsSmoker", SqlDbType.VarChar);	
						command.Parameters.Add("@ArEmpName", SqlDbType.NVarChar);	
						command.Parameters.Add("@ArFatherName", SqlDbType.NVarChar);	
						command.Parameters.Add("@ArLastName", SqlDbType.NVarChar);	
						command.Parameters.Add("@SpouseCharge", SqlDbType.VarChar);	
						command.Parameters.Add("@NoFaBen", SqlDbType.VarChar);	
						command.Parameters.Add("@Default", SqlDbType.VarChar);	
						command.Parameters.Add("@HoldFamilyAlloc", SqlDbType.VarChar);
						command.Parameters.Add("@EOSWithdrawDt", SqlDbType.NVarChar);
						command.Parameters.Add("@BasicPayUnit", SqlDbType.VarChar);
						command.Parameters.Add("@PayFrequency", SqlDbType.VarChar);
						command.Parameters.Add("@ValuePerUnit", SqlDbType.Decimal);
						command.Parameters.Add("@CurrID", SqlDbType.SmallInt);
						command.Parameters.Add("@TaxPolID", SqlDbType.Int);
						command.Parameters.Add("@PayPolicyID", SqlDbType.Int);
						command.Parameters.Add("@CurrRoundID", SqlDbType.Int);
						command.Parameters.Add("@UseClassDef", SqlDbType.Bit);
						command.Parameters.Add("@Spouse", SqlDbType.NVarChar);
						command.Parameters.Add("@SpName", SqlDbType.NVarChar);
						command.Parameters.Add("@SpLName", SqlDbType.NVarChar);
						command.Parameters.Add("@SpFaName", SqlDbType.NVarChar);
						command.Parameters.Add("@SpMoNameL", SqlDbType.NVarChar);
						command.Parameters.Add("@ArNationality", SqlDbType.NVarChar);
						command.Parameters.Add("@SpNationality", SqlDbType.NVarChar);
						command.Parameters.Add("@SpPOB", SqlDbType.NVarChar);
						command.Parameters.Add("@SpDOB", SqlDbType.NVarChar);
						command.Parameters.Add("@HireDate", SqlDbType.NVarChar);
						command.Parameters.Add("@TaxSince", SqlDbType.NVarChar);
						command.Parameters.Add("@HireSince", SqlDbType.NVarChar);
						command.Parameters.Add("@AccountNum", SqlDbType.NVarChar);
						///////////////////////////////////
						command.Parameters.Add("@CustomCr1", SqlDbType.Int);
						command.Parameters.Add("@CustomCr2", SqlDbType.Int);
						command.Parameters.Add("@Address1", SqlDbType.NVarChar);
						command.Parameters.Add("@Address2", SqlDbType.NVarChar);
						command.Parameters.Add("@Phone", SqlDbType.NVarChar);
						command.Parameters.Add("@Mobile", SqlDbType.NVarChar);
						command.Parameters.Add("@DisDate", SqlDbType.NVarChar);
						command.Parameters.Add("@DisCode", SqlDbType.Int);
						command.Parameters.Add("@MothName", SqlDbType.NVarChar);
						command.Parameters.Add("@Notes", SqlDbType.NVarChar);
						command.Parameters.Add("@Blood", SqlDbType.NVarChar);
						command.Parameters.Add("@ArMohafaza", SqlDbType.NVarChar);
						command.Parameters.Add("@ArKadaa", SqlDbType.NVarChar);
						command.Parameters.Add("@ArRegionTown", SqlDbType.NVarChar);
						command.Parameters.Add("@ArNeighborhood", SqlDbType.NVarChar);
						command.Parameters.Add("@ArStreet", SqlDbType.NVarChar);
						command.Parameters.Add("@ArEstateRegion", SqlDbType.NVarChar);
						command.Parameters.Add("@ArEstateNum", SqlDbType.NVarChar);
						command.Parameters.Add("@ArBuilding", SqlDbType.NVarChar);
						command.Parameters.Add("@ArFloor", SqlDbType.Int);
						command.Parameters.Add("@ArPhone1", SqlDbType.NVarChar);
						command.Parameters.Add("@ArPhone2", SqlDbType.NVarChar);
						command.Parameters.Add("@TAPolID", SqlDbType.Int);
						//======================
						//======================NEW
						//======================
						command.Parameters.Add("@CustomCr3", SqlDbType.NVarChar);
						command.Parameters.Add("@CustomCr4", SqlDbType.NVarChar);

						command.Parameters.Add("@POB", SqlDbType.NVarChar);
						command.Parameters.Add("@PassportNum", SqlDbType.NVarChar);
						command.Parameters.Add("@HighCostL", SqlDbType.Float);
						command.Parameters.Add("@IBAN", SqlDbType.NVarChar);

						command.Parameters.Add("@email", SqlDbType.NVarChar);
//						command.Parameters.Add("@isStudent", SqlDbType.Bit);
//						command.Parameters.Add("@isSmoker", SqlDbType.Bit);

						command.Parameters.Add("@ArMotherName", SqlDbType.NVarChar);
						command.Parameters.Add("@ArIDNum", SqlDbType.Int);
						command.Parameters.Add("@ArSijilKazaa", SqlDbType.NVarChar);
						command.Parameters.Add("@ArSijilMohafaza", SqlDbType.NVarChar);
						command.Parameters.Add("@ArSijilPlace", SqlDbType.NVarChar);

						command.Parameters.Add("@SpIDNum", SqlDbType.Int);
						command.Parameters.Add("@ArSijilNum", SqlDbType.NVarChar);
						command.Parameters.Add("@ArNationalityRegisterP", SqlDbType.NVarChar);
						
						command.Parameters.Add("@MaritalStatDate" , SqlDbType.VarChar);
						
						//						command.Parameters.Add("@emID", SqlDbType.Int);
						command.Parameters.Add("@BranchID", SqlDbType.Int);
						//======================
						//======================

						cmdEmpFamily.Parameters.Add("@EmpID", SqlDbType.Int);
						cmdEmpFamily.Parameters.Add("@Name", SqlDbType.NVarChar);
						cmdEmpFamily.Parameters.Add("@DOB", SqlDbType.NVarChar);
						cmdEmpFamily.Parameters.Add("@Gender", SqlDbType.VarChar);
						cmdEmpFamily.Parameters.Add("@OnCharge", SqlDbType.VarChar);
						cmdEmpFamily.Parameters.Add("@ChExemEndDate", SqlDbType.NVarChar);

						cmdPayItems.Parameters.Add("@ItmType", SqlDbType.VarChar);					
						cmdPayItems.Parameters.Add("@ItmDesc", SqlDbType.VarChar);
						cmdPayItems.Parameters.Add("@ItmAbb", SqlDbType.VarChar);
						cmdPayItems.Parameters.Add("@ItmCalc", SqlDbType.Int);
						cmdPayItems.Parameters.Add("@Rate", SqlDbType.Decimal);
						cmdPayItems.Parameters.Add("@CurrID", SqlDbType.SmallInt);
						cmdPayItems.Parameters.Add("@GlobalSalary", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@Special", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@Taxable", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@CalcEOS", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@CalcFA", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@CalcSick", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@CalcXtraEOS", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@LeaveUpdate", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@Reloan", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@InEmpCost", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@Printable", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@DedTax", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@Inactive", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@YearlyPayVac", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@Loan", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@Provision", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@R6", SqlDbType.Int);
						cmdPayItems.Parameters.Add("@Overtime", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@BasicPay", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@Istaxable", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@HighCost", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@IsTransport", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@School", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@ExtraMonth", SqlDbType.Bit);
						cmdPayItems.Parameters.Add("@HighCostOfLiving", SqlDbType.Bit);

						string tempEmpName = "";

						for (int i = 0; i < gridView3.RowCount; i++)
						{
							tempEmpName = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[FirstName]).ToString().Trim() + ' ' + gridView3.GetRowCellValue(i,gridView3.Columns[FatherName]).ToString() + ' ' + gridView3.GetRowCellValue(i,gridView3.Columns[LastName]).ToString());
							try
							{
								command.Parameters["@Name"].Value = tempEmpName.Substring(0, 50);
							}
							catch(Exception ex)
							{
								command.Parameters["@Name"].Value = tempEmpName;
							}

							command.Parameters["@firstname"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[FirstName])).Trim();
							command.Parameters["@middlename"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[FatherName]));
							command.Parameters["@lastname"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[LastName]));
							command.Parameters["@code"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[EmpCode])).Trim();
							command.Parameters["@TaxNb"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Tax])).Trim();
							if (gridView3.GetRowCellValue(i,gridView3.Columns[SocSec]) != null && gridView3.GetRowCellValue(i,gridView3.Columns[SocSec]) != DBNull.Value && gridView3.GetRowCellValue(i,gridView3.Columns[SocSec]).ToString().Trim().Length > 2)
							{
								command.Parameters["@SSN"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[SocSec]).ToString().Trim().Substring(0, 2) + "/" + gridView3.GetRowCellValue(i,gridView3.Columns[SocSec]).ToString().Trim().Substring(2).Trim();
							}
							else
							{
								command.Parameters["@SSN"].Value = DBNull.Value;
							}
							switch (gridView3.GetRowCellValue(i, gridView3.Columns[EmpType]).ToString().Trim())
							{
								case "OWNER": 
								case "owner":
									command.Parameters["@EmpType"].Value = 1;
									break;
								case "PA":
								case "pa":
									command.Parameters["@EmpType"].Value = 2;
									break;
								case "CO":
								case "co":
									command.Parameters["@EmpType"].Value = 3;
									break;
								case "PT":
								case "pt":
									command.Parameters["@EmpType"].Value = 4;
									break;
								case "FULL TIME":
								case "full time":
								case "FT":
								case "ft":
									command.Parameters["@EmpType"].Value = 5;
									break;
							}
							/////////////////////////////////////////////
							///                        -----Department-------
							///                        
//							command.Parameters["@Department"].Value = 4;
							string sqlDepID = @"Select [ID] from tbDepartment where Abbreviation = '" + gridView3.GetRowCellValue(i, gridView3.Columns[DepCode]).ToString().Trim() + "'";
							SqlCommand cmdDepID = new SqlCommand(sqlDepID, connection);
							connection.Open();
							object id1 = cmdDepID.ExecuteScalar();
							if (id1 == null || id1 == DBNull.Value)
							{
								cmDep.Parameters["@Desc"].Value = gridView3.GetRowCellValue(i, gridView3.Columns[DepDesc]).ToString();
								cmDep.Parameters["@DepCode"].Value = Convert.ToString(gridView3.GetRowCellValue(i, gridView3.Columns[DepCode])).Trim();	
								cmDep.ExecuteNonQuery();
								id1 = cmdDepID.ExecuteScalar();
								command.Parameters["@Department"].Value = id1;
							}
							else
							{
								command.Parameters["@Department"].Value = id1;
							}
							connection.Close();
							//======================

							//======================Branch
							string sqlBranchID = @"SELECT [ID] FROM tbBranch WHERE Abbreviation = '" + gridView3.GetRowCellValue(i, gridView3.Columns[BranchCode]).ToString().Trim() + "'";
							SqlCommand cmdBranchID = new SqlCommand(sqlBranchID, connection);
							connection.Open();
							object id4 = cmdBranchID.ExecuteScalar();
							if (id4 == null || id4 == DBNull.Value)
							{
								cmdBranch.Parameters["@BranchDesc"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[BranchDesc])).Trim();
								cmdBranch.Parameters["@BranchCode"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[BranchCode])).Trim();
								cmdBranch.ExecuteNonQuery();
								id4 = cmdBranchID.ExecuteScalar();
								command.Parameters["@BranchID"].Value = id4;
							}
							else
							{
								command.Parameters["@BranchID"].Value = id4;
							}
							connection.Close();
							//======================

							//======================Position
//							command.Parameters["@PositionA"].Value = 35;
							string sqlPosID = @"SELECT top 1 [ID] FROM tbPosition WHERE Abbreviation = '" + gridView3.GetRowCellValue(i, gridView3.Columns[PosCode]).ToString().Trim() + "'";
							SqlCommand cmdPosID = new SqlCommand(sqlPosID, connection);
							connection.Open();
							object id2 = cmdPosID.ExecuteScalar();
							if (id2 == null || id2 == DBNull.Value)
							{
								cmdPos.Parameters["@PosDesc"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[PosDesc])).Trim();
								cmdPos.Parameters["@PosCode"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[PosCode])).Trim();
								cmdPos.Parameters["@ArabicTitle"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[PosDescArabic]);
								cmdPos.ExecuteNonQuery();
								id2 = cmdPosID.ExecuteScalar();
								command.Parameters["@PositionA"].Value = id2;
							}
							else
							{
								command.Parameters["@PositionA"].Value = id2;
							}
							connection.Close();
							///                        -----Nationality-------
							string sqlNatID = "select [ID] from tbNationality where Abbreviation = '" +gridView3.GetRowCellValue(i,gridView3.Columns[NatCode]).ToString().Trim() + "'";
							SqlCommand cmdNatID = new SqlCommand(sqlNatID, connection);
							connection.Open();
							object id3 = cmdNatID.ExecuteScalar();
							if (id3 == null || id3 == DBNull.Value)
							{
								cmdNat.Parameters["@NatDesc"].Value = Convert.ToString(gridView3.GetRowCellValue(i, gridView3.Columns[NatDesc])).Trim();
								cmdNat.Parameters["@NatCode"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[NatCode])).Trim();
								if (NationalityDescArabic == "")
								{
									cmdNat.Parameters["@ArabicTitle"].Value = DBNull.Value;
								}
								else
								{
									cmdNat.Parameters["@ArabicTitle"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[NationalityDescArabic]);
								}
								cmdNat.ExecuteNonQuery();
								id3 = cmdNatID.ExecuteScalar();
								command.Parameters["@Nationality1"].Value = id3;
							}
							else
							{
								command.Parameters["@Nationality1"].Value = id3;
							}
							connection.Close();
							//-----------------------gridView3.GetRowCellValue(i, gridView3.Columns[DOB]).ToString() == ""
							if(gridView3.GetRowCellValue(i, gridView3.Columns[DOB]) == "")
							{
								command.Parameters["@DOB"].Value = "1980-01-01";
							}
							else
							{
								command.Parameters["@DOB"].Value = (gridView3.GetRowCellValue(i,gridView3.Columns[DOB]));
							}
							switch(gridView3.GetRowCellValue(i, gridView3.Columns[MaritalStatus]).ToString().Trim())
							{
								case "s":
								case "S":
								case "c":
								case "C":
									command.Parameters["@MarStat"].Value = "S";;
									break;
								case "m":
								case "M":
									command.Parameters["@MarStat"].Value = "M";;
									break;
								case "d":
								case "D":
									command.Parameters["@MarStat"].Value = "D";;
									break;
								case "w":
								case "W":
								case "v":
								case "V":
									command.Parameters["@MarStat"].Value = "W";;
									break;
							}
							switch (gridView3.GetRowCellValue(i, gridView3.Columns[Gender]).ToString().Trim())
							{
								case "F":
								case "f":
									command.Parameters["@Gender"].Value = "F";
									break;
								case "m":
								case "M":
									command.Parameters["@Gender"].Value = "M";
									break;
							}
							//------------------
							string sqlGenPol = "select max([ID]) from GenPolicy";
							SqlCommand cmdGenPol = new SqlCommand(sqlGenPol, connection);
							connection.Open();
							object idGen = cmdGenPol.ExecuteScalar();
							command.Parameters["@GenPolID"].Value = idGen;
							connection.Close();
							//-------------------
							string sqlTaxPol = "select max([ID]) from TaxPolicy";
							SqlCommand cmdTaxPol = new SqlCommand(sqlTaxPol,connection);
							connection.Open();
							object idTax = cmdTaxPol.ExecuteScalar();
							command.Parameters["@TaxPolID"].Value = idTax;
							connection.Close();
							//------------------
							string sqlTAPol = "select max([ID]) from ta_GenPol";
							SqlCommand cmdTAPol = new SqlCommand(sqlTAPol, connection);
							connection.Open();
							object idTA = cmdTAPol.ExecuteScalar();
							command.Parameters["@TAPolID"].Value = idTA;
							//command.Parameters["@TAPolID"].Value = DBNull.Value;
							connection.Close();
							//------------- determine pay policy
//							string tmpPM = "";
//							try
//							{
//								tmpPM = gridView3.GetRowCellValue(i, gridView3.Columns[PaymentMethod]).ToString();
//							}
//							catch(Exception Ex)
//							{
//								tmpPM = Ex.Message;
//							}

							if (gridView3.GetRowCellValue(i, gridView3.Columns[PaymentMethod]) == null || gridView3.GetRowCellValue(i, gridView3.Columns[PaymentMethod]) == DBNull.Value || gridView3.GetRowCellValue(i, gridView3.Columns[PaymentMethod]).ToString().Trim() == "" || (gridView3.GetRowCellValue(i, gridView3.Columns[PaymentMethod]).ToString().Trim() != "CH" && gridView3.GetRowCellValue(i, gridView3.Columns[PaymentMethod]).ToString().Trim() != "BA" ))
							{
								string sqlPayPol = "Select [ID] from tbPayPolicy where Title = 'Cash'";
								SqlCommand cmdPayPol = new SqlCommand(sqlPayPol, connection);
								connection.Open();
								object idPay = cmdPayPol.ExecuteScalar();
								if (idPay == null || idPay == DBNull.Value)
								{
									cmdPayPolicy.Parameters["@title"].Value = "Cash";
									cmdPayPolicy.Parameters["@regpaymethod"].Value = "C";
									cmdPayPolicy.Parameters["@specialpaymethod"].Value = "C";
									cmdPayPolicy.Parameters["@frombankidregular"].Value = DBNull.Value;
									cmdPayPolicy.Parameters["@frombankidspecial"].Value = DBNull.Value;
									cmdPayPolicy.ExecuteNonQuery();
									idPay = cmdPayPol.ExecuteScalar();
								}
								connection.Close();
								command.Parameters["@PayPolicyID"].Value = idPay;
							}
							else
							{
								if(gridView3.GetRowCellValue(i, gridView3.Columns[PaymentMethod]).ToString().Trim() == "BA")
								{
									SqlCommand cmdPayPol = new SqlCommand("select id from tbPayPolicy where title = 'BT-" + (gridView3.GetRowCellValue(i, gridView3.Columns[Bank]) == null || gridView3.GetRowCellValue(i, gridView3.Columns[Bank]) == DBNull.Value ? "'" : gridView3.GetRowCellValue(i, gridView3.Columns[Bank]).ToString().Trim() + "'"), connection);
									connection.Open();
									object idPay = cmdPayPol.ExecuteScalar();
									if (idPay == null || idPay == DBNull.Value)
									{
										cmdBank.Parameters["@bank"].Value = gridView3.GetRowCellValue(i, gridView3.Columns[Bank]);
										cmdBank.Parameters["@applytransfer"].Value = 1;
										cmdBank.ExecuteNonQuery();
										cmdPayPol.CommandText = "select max(id) from tbbank";
										idPay = cmdPayPol.ExecuteScalar();
										cmdBankSymbols.Parameters["@bankid"].Value = idPay;
										cmdBankSymbols.Parameters["@ourcurrency"].Value = 149;
										cmdBankSymbols.Parameters["@banksymbol"].Value = "LBP";
										cmdBankSymbols.ExecuteNonQuery();
										cmdBankSymbols.Parameters["@ourcurrency"].Value = 150;
										cmdBankSymbols.Parameters["@banksymbol"].Value = "USD";
										cmdBankSymbols.ExecuteNonQuery();
										if (gridView3.GetRowCellValue(i,gridView3.Columns[CyOfSalary]).ToString() == "LBP")
										{
											cmdBankAccount.Parameters["@currid"].Value = 149;
										}
										else if (gridView3.GetRowCellValue(i,gridView3.Columns[CyOfSalary]).ToString() == "USD")
										{
											cmdBankAccount.Parameters["@currid"].Value = 150;
										}
										cmdBankAccount.Parameters["@bankid"].Value = idPay;
										cmdBankAccount.Parameters["@bankaccount"].Value = "0";
										cmdBankAccount.Parameters["@accountdesc"].Value = (gridView3.GetRowCellValue(i, gridView3.Columns[Bank]) == null || gridView3.GetRowCellValue(i, gridView3.Columns[Bank]) == DBNull.Value ? "" : gridView3.GetRowCellValue(i, gridView3.Columns[Bank]).ToString().Trim()) + " Account";
										cmdBankAccount.ExecuteNonQuery();
										cmdPayPol.CommandText = "select max(id) from tbbankaccounts";
										idPay = cmdPayPol.ExecuteScalar();
										cmdPayPolicy.Parameters["@title"].Value = "BT-" + (gridView3.GetRowCellValue(i, gridView3.Columns[Bank]) == null || gridView3.GetRowCellValue(i, gridView3.Columns[Bank]) == DBNull.Value ? "" : gridView3.GetRowCellValue(i, gridView3.Columns[Bank]).ToString().Trim());
										cmdPayPolicy.Parameters["@regpaymethod"].Value = "BA";
										cmdPayPolicy.Parameters["@specialpaymethod"].Value = "BA";
										cmdPayPolicy.Parameters["@frombankidregular"].Value = idPay;
										cmdPayPolicy.Parameters["@frombankidspecial"].Value = idPay;
										cmdPayPolicy.ExecuteNonQuery();
										cmdPayPol.CommandText = "select id from tbPayPolicy where title = 'BT-" + (gridView3.GetRowCellValue(i, gridView3.Columns[Bank]) == null || gridView3.GetRowCellValue(i, gridView3.Columns[Bank]) == DBNull.Value ? "'" : gridView3.GetRowCellValue(i, gridView3.Columns[Bank]).ToString().Trim() + "'");
										idPay = cmdPayPol.ExecuteScalar();

									}
									command.Parameters["@PayPolicyID"].Value = idPay;
									connection.Close();
								}
								else
								{
									//==================
									//=========CH: Check
									string sqlPayPol = "Select [ID] from tbPayPolicy where Title = 'Cheque'";
									SqlCommand cmdPayPol = new SqlCommand(sqlPayPol, connection);
									connection.Open();
									object idPay = cmdPayPol.ExecuteScalar();
									if (idPay == null || idPay == DBNull.Value)
									{
										cmdPayPolicy.Parameters["@title"].Value = "Cheque";
										cmdPayPolicy.Parameters["@regpaymethod"].Value = "C";
										cmdPayPolicy.Parameters["@specialpaymethod"].Value = "C";
										cmdPayPolicy.Parameters["@frombankidregular"].Value = DBNull.Value;
										cmdPayPolicy.Parameters["@frombankidspecial"].Value = DBNull.Value;
										cmdPayPolicy.ExecuteNonQuery();
										idPay = cmdPayPol.ExecuteScalar();
									}
									connection.Close();
									command.Parameters["@PayPolicyID"].Value = idPay;
									//==================
								}
							}
							//----------------------
							string sqlCurr = "select max( [ID]) from tbCurrRoundPolicies";
							SqlCommand cmdCurr = new SqlCommand(sqlCurr,connection);
							connection.Open();
							object idCurr = cmdCurr.ExecuteScalar();
							command.Parameters["@CurrRoundID"].Value = idCurr;
							connection.Close();
							//----------------------
							command.Parameters["@EOSWithdrawDt"].Value = DBNull.Value; // (gridView3.GetRowCellValue(i,gridView3.Columns[EOS]));
							switch(gridView3.GetRowCellValue(i,gridView3.Columns[BasicSalaryUnit]).ToString().Trim())
							{
								case "m":
								case "M":
									command.Parameters["@BasicPayUnit"].Value = "M";;
									break;
								case "d":
								case "D":
									command.Parameters["@BasicPayUnit"].Value = "D";;
									break;
								case "h":
								case "H":
									command.Parameters["@BasicPayUnit"].Value = "H";;
									break;
							}
							switch(gridView3.GetRowCellValue(i,gridView3.Columns[PayFrequency]).ToString().Trim())
							{
								case "m":
								case "M":
									command.Parameters["@PayFrequency"].Value = "M";;
									break;
								case "w":
								case "W":
									command.Parameters["@PayFrequency"].Value = "W";;
									break;
								case "bw":
								case "BW":
									command.Parameters["@PayFrequency"].Value = "BW";;
									break;
								case "sm":
								case "SM":
									command.Parameters["@PayFrequency"].Value = "SM";;
									break;
							}
							command.Parameters["@ValuePerUnit"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[SalaryValue]);
							if (gridView3.GetRowCellValue(i,gridView3.Columns[CyOfSalary]).ToString() == "LBP")
							{
								command.Parameters["@CurrID"].Value = 149;
							}
							else
								if (gridView3.GetRowCellValue(i,gridView3.Columns[CyOfSalary]).ToString() == "USD")
							{
								command.Parameters["@CurrID"].Value = 150;
							}
							else
							{
								command.Parameters["@CurrID"].Value = DBNull.Value;
							}
							//=============================
							//==============check if saved
							//=============================
							command.Parameters["@RegNum"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[Reg]);
							command.Parameters["@BadgeNum"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[TimeAtt])).Trim();
							//=============================
							//=============================

							command.Parameters["@Inactive"].Value = "Ac";
							if (gridView3.GetRowCellValue(i,gridView3.Columns[SpouseWork]) != null && gridView3.GetRowCellValue(i,gridView3.Columns[SpouseWork]) != DBNull.Value && gridView3.GetRowCellValue(i,gridView3.Columns[SpouseWork]).ToString().Trim().ToUpper() == "Y")
							{
								command.Parameters["@SpouseWork"].Value = 1;
							}
							else
							{
								command.Parameters["@SpouseWork"].Value = 0;
							}
							command.Parameters["@Decease"].Value = 0;
							command.Parameters["@IsStudent"].Value = 0;
							command.Parameters["@IsSmoker"].Value = 0;
							command.Parameters["@ArEmpName"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[FirstNameArabic]);
							command.Parameters["@ArFatherName"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[FatherNameArabic]);
							command.Parameters["@ArLastName"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[LastNameArabic]);
							command.Parameters["@SpouseCharge"].Value = 0;
							command.Parameters["@NoFaBen"].Value = 0;
							command.Parameters["@Default"].Value = 0;
							command.Parameters["@HoldFamilyAlloc"].Value = 0;
							command.Parameters["@UseClassDef"].Value = 0;
							command.Parameters["@HireDate"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[HiringDate]);
							//=============================
							//==============check if saved
							//=============================
							if (SpFirstName == "")
							{
								command.Parameters["@SpName"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@SpName"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[SpFirstName]);
							}

							if (SpLastName == "")
							{
								command.Parameters["@SpLName"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@SpLName"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[SpLastName]);
							}
							if (SpFatherName == "")
							{
								command.Parameters["@SpFaName"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@SpFaName"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[SpFatherName]);
							}
							command.Parameters["@SpMoNameL"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[SpMotherName]);
							command.Parameters["@ArNationality"].Value = DBNull.Value;
							if (SpNationality == "")
							{
								command.Parameters["@SpNationality"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@SpNationality"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[SpNationality]);
							}
							if (SpPOB == "")
							{
								command.Parameters["@SpPOB"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@SpPOB"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[SpPOB]);
							}
							if (SpDOB == "")
							{
								command.Parameters["@SpDOB"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@SpDOB"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[SpDOB]);
							}
							command.Parameters["@AccountNum"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[BankAccount]);
							
//							command.Parameters["@SpName"].Value = DBNull.Value;
//							command.Parameters["@SpLName"].Value = DBNull.Value;
//							command.Parameters["@SpFaName"].Value = DBNull.Value;
//							command.Parameters["@SpNationality"].Value = DBNull.Value;
//							command.Parameters["@SpPOB"].Value = DBNull.Value;
//							command.Parameters["@SpDOB"].Value = DBNull.Value;
//							command.Parameters["@SpMoNameL"].Value = DBNull.Value;

							//------------------------------ new additions
							if (Spouse == "")
							{
								command.Parameters["@Spouse"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@Spouse"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[Spouse]);
							}
							if (DisDate == "" || gridView3.GetRowCellValue(i, gridView3.Columns[DisDate]) == null || gridView3.GetRowCellValue(i, gridView3.Columns[DisDate]) == DBNull.Value || gridView3.GetRowCellValue(i, gridView3.Columns[DisDate]).ToString().Trim().Length == 0)
							{
								command.Parameters["@Discharge"].Value = 0;
								command.Parameters["@DisDate"].Value = DBNull.Value;
								command.Parameters["@DisCode"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@Discharge"].Value = 1;
								command.Parameters["@DisDate"].Value = gridView3.GetRowCellValue(i, gridView3.Columns[DisDate]);
								command.Parameters["@DisCode"].Value = 1;
							}
							//=============================
							//=============================
							if (TaxSince == "")
							{
								command.Parameters["@TaxSince"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[HiringDate]);
							}
							else
							{
								if (gridView3.GetRowCellValue(i,gridView3.Columns[TaxSince]) == DBNull.Value || gridView3.GetRowCellValue(i,gridView3.Columns[TaxSince]).ToString().Trim() == "")
								{
									command.Parameters["@TaxSince"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[HiringDate]);
								}
								else
								{
									command.Parameters["@TaxSince"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[TaxSince]);
								}
							}
							if (NSSFSince == "")
							{
								command.Parameters["@HireSince"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[HiringDate]);
							}
							else
							{
								if (gridView3.GetRowCellValue(i,gridView3.Columns[NSSFSince]) == DBNull.Value || gridView3.GetRowCellValue(i,gridView3.Columns[NSSFSince]).ToString().Trim() == "")
								{
									command.Parameters["@HireSince"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[HiringDate]);
								}
								else
								{
									command.Parameters["@HireSince"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[NSSFSince]);
								}
							}
							//------------------------ CustomCr1
//							command.Parameters["@CustomCr1"].Value = DBNull.Value;
							if (Custom1Desc == "" || this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Custom1Desc]) == null || this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Custom1Desc]) == DBNull.Value || this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Custom1Desc]).ToString().Trim().Length == 0)
							{
								command.Parameters["@CustomCr1"].Value = DBNull.Value;
							}
							else
							{
								string sqlCustom1ID = @"SELECT [ID] FROM tbCustomCr1 WHERE Abbreviation = '" + gridView3.GetRowCellValue(i, gridView3.Columns[Custom1Code]).ToString().Trim() + "'";
								SqlCommand cmdCustom1ID = new SqlCommand(sqlCustom1ID, connection);
								connection.Open();
								object id = cmdCustom1ID.ExecuteScalar();
								if (id == null || id == DBNull.Value)
								{
									cmdCustom1.Parameters["@Title"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Custom1Desc])).Trim();
									if (Custom1Code == "")
									{
										cmdCustom1.Parameters["@Abbreviation"].Value = DBNull.Value;
									}
									else
									{
										cmdCustom1.Parameters["@Abbreviation"].Value = this.gridView3.GetRowCellValue(i,gridView3.Columns[Custom1Code]);
									}
									cmdCustom1.ExecuteNonQuery();
									id = cmdCustom1ID.ExecuteScalar();
									command.Parameters["@CustomCr1"].Value = id;
								}
								else
								{
									command.Parameters["@CustomCr1"].Value = id;
								}
								connection.Close();
							}
							//------------------------ CustomCr2
							command.Parameters["@CustomCr2"].Value = DBNull.Value;
							if (Custom2Desc == "" || this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Custom2Desc]) == null || this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Custom2Desc]) == DBNull.Value || this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Custom2Desc]).ToString().Trim().Length == 0)
							{
								command.Parameters["@CustomCr2"].Value = DBNull.Value;
							}
							else
							{
								string sqlCustom2ID = @"SELECT [ID] FROM tbCustomCr2 WHERE Title = '" + gridView3.GetRowCellValue(i, gridView3.Columns[Custom2Desc]).ToString().Trim() + "'";
								SqlCommand cmdCustom2ID = new SqlCommand(sqlCustom2ID, connection);
								connection.Open();
								object id = cmdCustom2ID.ExecuteScalar();
								if (id == null || id == DBNull.Value)
								{
									cmdCustom2.Parameters["@Title"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Custom2Desc])).Trim();
									if (Custom2Code == "")
									{
										cmdCustom2.Parameters["@Abbreviation"].Value = DBNull.Value;
									}
									else
									{
										cmdCustom2.Parameters["@Abbreviation"].Value = this.gridView3.GetRowCellValue(i,gridView3.Columns[Custom2Code]);
									}
									cmdCustom2.ExecuteNonQuery();
									id = cmdCustom2ID.ExecuteScalar();
									command.Parameters["@CustomCr2"].Value = id;
								}
								else
								{
									command.Parameters["@CustomCr2"].Value = id;
								}
								connection.Close();
							}
							//---------------------
							if (Address1 == "")
							{
								command.Parameters["@Address1"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@Address1"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Address1]);
							}
							if (Address2 == "")
							{
								command.Parameters["@Address2"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@Address2"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Address2]);
							}
							if (Phone == "")
							{
								command.Parameters["@Phone"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@Phone"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Phone]);
							}
							if (Mobile == "")
							{
								command.Parameters["@Mobile"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@Mobile"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Mobile]);
							}
							if (MothName == "")
							{
								command.Parameters["@MothName"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@MothName"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[MothName]);
							}
							if (Notes == "")
							{
								command.Parameters["@Notes"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@Notes"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Notes]);
							}
							if (Blood == "")
							{
								command.Parameters["@Blood"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@Blood"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Blood]);
							}
							if (ArMohafaza == "")
							{
								command.Parameters["@ArMohafaza"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@ArMohafaza"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArMohafaza]);
							}
							if (ArKadaa == "")
							{
								command.Parameters["@ArKadaa"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@ArKadaa"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArKadaa]);
							}
							if (ArRegionTown == "")
							{
								command.Parameters["@ArRegionTown"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@ArRegionTown"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArRegionTown]);
							}
							if (ArNeighborhood == "")
							{
								command.Parameters["@ArNeighborhood"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@ArNeighborhood"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArNeighborhood]);
							}
							if (ArStreet == "")
							{
								command.Parameters["@ArStreet"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@ArStreet"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArStreet]);
							}
							if (ArEstateRegion == "")
							{
								command.Parameters["@ArEstateRegion"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@ArEstateRegion"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArEstateRegion]);
							}
							if (ArEstateNum == "")
							{
								command.Parameters["@ArEstateNum"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@ArEstateNum"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArEstateNum]);
							}
							if (ArBuilding == "")
							{
								command.Parameters["@ArBuilding"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@ArBuilding"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArBuilding]);
							}

							command.Parameters["@ArFloor"].Value = DBNull.Value;
//							if (ArFloor == "")
//							{
//								command.Parameters["@ArFloor"].Value = DBNull.Value;
//							}
//							else
//							{
//								command.Parameters["@ArFloor"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArFloor]);
//							}
							if (ArPhone1 == "")
							{
								command.Parameters["@ArPhone1"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@ArPhone1"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArPhone1]);
							}
							if (ArPhone2 == "")
							{
								command.Parameters["@ArPhone2"].Value = DBNull.Value;
							}
							else
							{
								command.Parameters["@ArPhone2"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArPhone2]);
							}
							//---------------------
							
							//======================
							//======================NEW
							//======================
							command.Parameters["@CustomCr3"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[CustomCr3]);
							command.Parameters["@CustomCr4"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[CustomCr4]);

							command.Parameters["@POB"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[PlaceOfBirth]);
							command.Parameters["@PassportNum"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[PassportNb]);
							command.Parameters["@HighCostL"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[highCostOfLiving]);
							command.Parameters["@IBAN"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[IBAN]);

							command.Parameters["@email"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[Email]);
							command.Parameters["@isStudent"].Value = 0;//this.gridView3.GetRowCellValue(i, this.gridView3.Columns[IsStudent]);
							command.Parameters["@isSmoker"].Value = 0;//this.gridView3.GetRowCellValue(i, this.gridView3.Columns[IsSmoker]);

							command.Parameters["@ArMotherName"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArMotherName]);
							command.Parameters["@ArIDNum"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[IdCardNb]);
							command.Parameters["@ArSijilKazaa"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArKazaCard]);
							command.Parameters["@ArSijilMohafaza"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArMohafazaCard]);
							command.Parameters["@ArSijilPlace"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[ArRegisterPlace]);
						
							command.Parameters["@SpIDNum"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[SpIdCardNb]);
							command.Parameters["@ArSijilNum"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[SpRegisterNb]);
							command.Parameters["@ArNationalityRegisterP"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[SpRegisterPlace]);
							
							command.Parameters["@MaritalStatDate"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[DOB]); //this.gridView3.GetRowCellValue(i, this.gridView3.Columns[MaritalStatusDate]);
							//======================
							//======================
							connection.Open();  
							command.ExecuteNonQuery();
							connection.Close();

////////////////////////////////////////////////////////////////////////////////////////////////
///Child
///Transaportation
///////////////////////
							if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1Name]).ToString() != "")
							{
								string sqlEmpID = "Select max ([ID]) from Emp";
							
								SqlCommand cmdChild = new SqlCommand(sqlEmpID,connection);
								connection.Open();
								int id = Convert.ToInt32(cmdChild.ExecuteScalar());
								cmdEmpFamily.Parameters["@EmpID"].Value = id;
								cmdEmpFamily.Parameters["@Name"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child1Name]));
								cmdEmpFamily.Parameters["@DOB"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child1DOB]));
								cmdEmpFamily.Parameters["@Gender"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child1Gender])).Substring(0, 1).ToUpper();
								cmdEmpFamily.Parameters["@OnCharge"].Value = "Until Age Limit";
								if (gridView3.GetRowCellValue(i,gridView3.Columns[Child1TillDate]) == null || gridView3.GetRowCellValue(i,gridView3.Columns[Child1TillDate]) == DBNull.Value || gridView3.GetRowCellValue(i,gridView3.Columns[Child1TillDate]).ToString().Trim() == "")
								{
									cmdEmpFamily.Parameters["@ChExemEndDate"].Value = DBNull.Value;
								}
								else
								{
									cmdEmpFamily.Parameters["@ChExemEndDate"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child1TillDate]));
								}
								cmdEmpFamily.ExecuteNonQuery();
								connection.Close();
							}
							if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2Name]).ToString() != "")
							{
								string sqlEmpID = "Select max ([ID]) from Emp ";
							
								SqlCommand cmdChild = new SqlCommand(sqlEmpID,connection);
								connection.Open();
								int id = Convert.ToInt32(cmdChild.ExecuteScalar());
								cmdEmpFamily.Parameters["@EmpID"].Value = id;
								cmdEmpFamily.Parameters["@Name"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child2Name]));
								cmdEmpFamily.Parameters["@DOB"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child2DOB]));
								cmdEmpFamily.Parameters["@Gender"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child2Gender])).Substring(0, 1).ToUpper();
								cmdEmpFamily.Parameters["@OnCharge"].Value = "Until Age Limit";
								if (gridView3.GetRowCellValue(i,gridView3.Columns[Child2TillDate]) == null || gridView3.GetRowCellValue(i,gridView3.Columns[Child2TillDate]) == DBNull.Value || gridView3.GetRowCellValue(i,gridView3.Columns[Child2TillDate]).ToString().Trim() == "")
								{
									cmdEmpFamily.Parameters["@ChExemEndDate"].Value = DBNull.Value;
								}
								else
								{
									cmdEmpFamily.Parameters["@ChExemEndDate"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child2TillDate]));
								}
								cmdEmpFamily.ExecuteNonQuery();
								connection.Close();
							}
							if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3Name]).ToString() != "")
							{
								string sqlEmpID = "Select max ([ID]) from Emp ";
							
								SqlCommand cmdChild = new SqlCommand(sqlEmpID,connection);
								connection.Open();
								int id = Convert.ToInt32(cmdChild.ExecuteScalar());
								cmdEmpFamily.Parameters["@EmpID"].Value = id;
								cmdEmpFamily.Parameters["@Name"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child3Name]));
								cmdEmpFamily.Parameters["@DOB"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child3DOB]));
								cmdEmpFamily.Parameters["@Gender"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child3Gender])).Substring(0, 1).ToUpper();
								cmdEmpFamily.Parameters["@OnCharge"].Value = "Until Age Limit";
								if (gridView3.GetRowCellValue(i,gridView3.Columns[Child3TillDate]) == null || gridView3.GetRowCellValue(i,gridView3.Columns[Child3TillDate]) == DBNull.Value || gridView3.GetRowCellValue(i,gridView3.Columns[Child3TillDate]).ToString().Trim() == "")
								{
									cmdEmpFamily.Parameters["@ChExemEndDate"].Value = DBNull.Value;
								}
								else
								{
									cmdEmpFamily.Parameters["@ChExemEndDate"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child3TillDate]));
								}
								cmdEmpFamily.ExecuteNonQuery();
								connection.Close();
							}
							if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4Name]).ToString() != "")
							{
								string sqlEmpID = "Select max ([ID]) from Emp";
							
								SqlCommand cmdChild = new SqlCommand(sqlEmpID,connection);
								connection.Open();
								int id = Convert.ToInt32(cmdChild.ExecuteScalar());
								cmdEmpFamily.Parameters["@EmpID"].Value = id;
								cmdEmpFamily.Parameters["@Name"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child4Name]));
								cmdEmpFamily.Parameters["@DOB"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child4DOB]));
								cmdEmpFamily.Parameters["@Gender"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child4Gender])).Substring(0, 1).ToUpper();
								cmdEmpFamily.Parameters["@OnCharge"].Value = "Until Age Limit";
								if (gridView3.GetRowCellValue(i,gridView3.Columns[Child4TillDate]) == null || gridView3.GetRowCellValue(i,gridView3.Columns[Child4TillDate]) == DBNull.Value || gridView3.GetRowCellValue(i,gridView3.Columns[Child4TillDate]).ToString().Trim() == "")
								{
									cmdEmpFamily.Parameters["@ChExemEndDate"].Value = DBNull.Value;
								}
								else
								{
									cmdEmpFamily.Parameters["@ChExemEndDate"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child4TillDate]));
								}
								cmdEmpFamily.ExecuteNonQuery();
								connection.Close();
							}
							if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5Name]).ToString() != "")
							{
								string sqlEmpID = "Select max ([ID]) from Emp ";
							
								SqlCommand cmdChild = new SqlCommand(sqlEmpID,connection);
								connection.Open();
								int id = Convert.ToInt32(cmdChild.ExecuteScalar());
								cmdEmpFamily.Parameters["@EmpID"].Value = id;
								cmdEmpFamily.Parameters["@Name"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child5Name]));
								cmdEmpFamily.Parameters["@DOB"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child5DOB]));
								cmdEmpFamily.Parameters["@Gender"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child5Gender])).Substring(0, 1).ToUpper();
								cmdEmpFamily.Parameters["@OnCharge"].Value = "Until Age Limit";
								if (gridView3.GetRowCellValue(i,gridView3.Columns[Child5TillDate]) == null || gridView3.GetRowCellValue(i,gridView3.Columns[Child5TillDate]) == DBNull.Value || gridView3.GetRowCellValue(i,gridView3.Columns[Child5TillDate]).ToString().Trim() == "")
								{
									cmdEmpFamily.Parameters["@ChExemEndDate"].Value = DBNull.Value;
								}
								else
								{
									cmdEmpFamily.Parameters["@ChExemEndDate"].Value = Convert.ToString(gridView3.GetRowCellValue(i,gridView3.Columns[Child5TillDate]));
								}
								cmdEmpFamily.ExecuteNonQuery();
								connection.Close();
							}

							if (gridView3.GetRowCellValue(i,gridView3.Columns[TranspDesc]) != DBNull.Value && gridView3.GetRowCellValue(i,gridView3.Columns[TranspDesc]).ToString().Trim() != "" && gridView3.GetRowCellValue(i,gridView3.Columns[TransValue]) != DBNull.Value && Convert.ToDecimal(gridView3.GetRowCellValue(i,gridView3.Columns[TransValue])) != 0)
							{

								string sqlPayDesc = "Select [ID] from PayItems where ItmDesc = '"+ gridView3.GetRowCellValue(i,gridView3.Columns[TranspDesc]).ToString()+ "'";
								SqlCommand cmdPayDesc = new SqlCommand(sqlPayDesc, connection);
								connection.Open();
								int idTr = Convert.ToInt32(cmdPayDesc.ExecuteScalar());
								if (idTr == 0)
								{
									cmdPayItems.Parameters["@ItmType"].Value = "I"; 
									cmdPayItems.Parameters["@ItmDesc"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[TranspDesc]).ToString();
									cmdPayItems.Parameters["@ItmAbb"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[TranspCode]).ToString();
									
									object test = gridView3.GetRowCellValue(i,gridView3.Columns[TransUnit]);

									if (gridView3.GetRowCellValue(i,gridView3.Columns[TransUnit]).ToString() == "D")
									{
										cmdPayItems.Parameters["@ItmCalc"].Value = 4;
									}
									else
									{
										cmdPayItems.Parameters["@ItmCalc"].Value = 3;
									}
									cmdPayItems.Parameters["@Rate"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[TransValue]);
									if (gridView3.GetRowCellValue(i,gridView3.Columns[CyOfSalary]).ToString() == "LBP")
									{
										cmdPayItems.Parameters["@CurrID"].Value = 149;
									}
									else
									{
										cmdPayItems.Parameters["@CurrID"].Value = 150;
									}
									cmdPayItems.Parameters["@GlobalSalary"].Value = 0;
									cmdPayItems.Parameters["@Special"].Value = 0;
									cmdPayItems.Parameters["@Taxable"].Value = 0;
									cmdPayItems.Parameters["@CalcEOS"].Value = 0;
									cmdPayItems.Parameters["@CalcFA"].Value = 0;
									cmdPayItems.Parameters["@CalcSick"].Value = 0;
									cmdPayItems.Parameters["@CalcXtraEOS"].Value =0; 
									cmdPayItems.Parameters["@LeaveUpdate"].Value = 0;
									cmdPayItems.Parameters["@Reloan"].Value = 0;
									cmdPayItems.Parameters["@InEmpCost"].Value =0; 
									cmdPayItems.Parameters["@Printable"].Value = 0;
									cmdPayItems.Parameters["@DedTax"].Value = 0;
									cmdPayItems.Parameters["@Inactive"].Value = 0;
									cmdPayItems.Parameters["@YearlyPayVac"].Value = 0;
									cmdPayItems.Parameters["@Loan"].Value = 0;
									cmdPayItems.Parameters["@Provision"].Value =0; 
									cmdPayItems.Parameters["@R6"].Value = 150;
									cmdPayItems.Parameters["@Overtime"].Value =0; 
									cmdPayItems.Parameters["@BasicPay"].Value = 0;
									cmdPayItems.Parameters["@Istaxable"].Value = 1;
									cmdPayItems.Parameters["@HighCost"].Value = 0;
									cmdPayItems.Parameters["@IsTransport"].Value = 1;
									cmdPayItems.Parameters["@School"].Value = 0;
									cmdPayItems.Parameters["@ExtraMonth"].Value = 0;
									cmdPayItems.Parameters["@HighCostOfLiving"].Value = 0;
									cmdPayItems.ExecuteNonQuery();
									idTr = Convert.ToInt32(cmdPayDesc.ExecuteScalar());
								}
								string sqlEmpID = "Select max ([ID]) from Emp";
								SqlCommand cmdChild = new SqlCommand(sqlEmpID,connection);
								int id = Convert.ToInt32(cmdChild.ExecuteScalar());
								cmdEmpPayItems.Parameters["@empid"].Value = id;
								cmdEmpPayItems.Parameters["@paytypeid"].Value = idTr;
								if (gridView3.GetRowCellValue(i,gridView3.Columns[CyOfSalary]).ToString() == "LBP")
								{
									cmdEmpPayItems.Parameters["@currid"].Value = 149;
								}
								else
								{
									cmdEmpPayItems.Parameters["@currid"].Value = 150;
								}
								
								cmdEmpPayItems.Parameters["@value"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[TransValue]);
								cmdEmpPayItems.Parameters["@paymentfreq"].Value = "EP";
								cmdEmpPayItems.Parameters["@type"].Value = "Normal";
								cmdEmpPayItems.ExecuteNonQuery();
								connection.Close();
							}
							
							//======================
							//========marital status
							//======================
							if (gridView3.GetRowCellValue(i,gridView3.Columns[MaritalStatus]) != DBNull.Value && gridView3.GetRowCellValue(i,gridView3.Columns[MaritalStatus]).ToString().Trim() != "S")
							{
								string sqlEmpIDms = "Select max ([ID]) from Emp";
								SqlCommand cmdMaritalStId = new SqlCommand(sqlEmpIDms,connection);
								connection.Open();
								int idem = Convert.ToInt32(cmdMaritalStId.ExecuteScalar());

								cmdMaritalStatus.Parameters["@emID"].Value = idem;
								cmdMaritalStatus.Parameters["@MaritalStatSymbol"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[MaritalStatus]);
								cmdMaritalStatus.Parameters["@MaritalStatDate"].Value = this.gridView3.GetRowCellValue(i, this.gridView3.Columns[MaritalStatusDate]);
								
								if (gridView3.GetRowCellValue(i,gridView3.Columns[SpouseWork]) != null && gridView3.GetRowCellValue(i,gridView3.Columns[SpouseWork]) != DBNull.Value && gridView3.GetRowCellValue(i,gridView3.Columns[SpouseWork]).ToString().Trim().ToUpper() == "Y")
								{
									cmdMaritalStatus.Parameters["@SpouseWork"].Value = 1;
								}
								else
								{
									cmdMaritalStatus.Parameters["@SpouseWork"].Value = 0;
								}
								
								if(gridView3.GetRowCellValue(i,gridView3.Columns[HeafOfFamily]) != null && gridView3.GetRowCellValue(i,gridView3.Columns[HeafOfFamily]) != DBNull.Value && gridView3.GetRowCellValue(i,gridView3.Columns[HeafOfFamily]).ToString().Trim().ToUpper() == "Y")
									cmdMaritalStatus.Parameters["@headoffamily"].Value = 1;
								else
									cmdMaritalStatus.Parameters["@headoffamily"].Value = 0;
								
								cmdMaritalStatus.ExecuteNonQuery();
								connection.Close();
							}
							
							//======================
//							if (gridView3.GetRowCellValue(i,gridView3.Columns[IncomePayItemValue1]) != DBNull.Value && gridView3.GetRowCellValue(i,gridView3.Columns[IncomePayItemValue1]).ToString().Trim() != "")
//							 {
//								 string sqlPI1 = "Select max ([ID]) from Emp";
//								 SqlCommand cmdPI1 = new SqlCommand(sqlPI1,connection);
//								 connection.Open();
//								 int idemPI1 = Convert.ToInt32(cmdPI1.ExecuteScalar());
//
//								 cmdPayItem1.Parameters["@emID1"].Value = idemPI1;
//								 cmdPayItem1.Parameters["@piValue1"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[IncomePayItemValue1]);
//								
//								 cmdPayItem1.ExecuteNonQuery();
//								 connection.Close();
//							 }
							//======================
//							if (gridView3.GetRowCellValue(i,gridView3.Columns[IncomePayItemValue2]) != DBNull.Value && gridView3.GetRowCellValue(i,gridView3.Columns[IncomePayItemValue2]).ToString().Trim() != "")
//							{
//								string sqlPI2 = "Select max ([ID]) from Emp";
//								SqlCommand cmdPI2 = new SqlCommand(sqlPI2,connection);
//								connection.Open();
//								int idemPI2 = Convert.ToInt32(cmdPI2.ExecuteScalar());
//
//								cmdPayItem2.Parameters["@emID2"].Value = idemPI2;
//								cmdPayItem2.Parameters["@piValue2"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[IncomePayItemValue2]);
//								
//								cmdPayItem2.ExecuteNonQuery();
//								connection.Close();
//							}
							//======================
//							if (gridView3.GetRowCellValue(i,gridView3.Columns[DeductPayItemValue1]) != DBNull.Value && gridView3.GetRowCellValue(i,gridView3.Columns[DeductPayItemValue1]).ToString().Trim() != "")
//							{
//								string sqlPID1 = "Select max ([ID]) from Emp";
//								SqlCommand cmdPID1 = new SqlCommand(sqlPID1,connection);
//								connection.Open();
//								int idemPID1 = Convert.ToInt32(cmdPID1.ExecuteScalar());
//
//								cmdPayItemDed1.Parameters["@emIDd1"].Value = idemPID1;
//								cmdPayItemDed1.Parameters["@pidValue1"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[DeductPayItemValue1]);
//								
//								cmdPayItemDed1.ExecuteNonQuery();
//								connection.Close();
//							}
							//======================
//							if (gridView3.GetRowCellValue(i,gridView3.Columns[DeductPayItemValue2]) != DBNull.Value && gridView3.GetRowCellValue(i,gridView3.Columns[DeductPayItemValue2]).ToString().Trim() != "")
//							{
//								string sqlPID2 = "Select max ([ID]) from Emp";
//								SqlCommand cmdPID2 = new SqlCommand(sqlPID2,connection);
//								connection.Open();
//								int idemPID2 = Convert.ToInt32(cmdPID2.ExecuteScalar());
//
//								cmdPayItemDed2.Parameters["@emIDd2"].Value = idemPID2;
//								cmdPayItemDed2.Parameters["@pidValue2"].Value = gridView3.GetRowCellValue(i,gridView3.Columns[DeductPayItemValue2]);
//								
//								cmdPayItemDed2.ExecuteNonQuery();
//								connection.Close();
//							}
							//======================
						}
						command.Parameters.Clear();
						connection.Open();
						command.CommandText = "insert into emphistory (empid, effdate, marstat, spousework, headoffamily, nocfaben) select id, dob, 'S', 0, 0, 0 from emp";
						command.ExecuteNonQuery();
//						command.CommandText = "insert into emphistory (empid, effdate, marstat, spousework, headoffamily, nocfaben) select id, dateadd(day, 1, dob), 'M', spousework, 0, 0 from emp where marstat = 'M'";
//						command.ExecuteNonQuery();
						MessageBox.Show("It's Done!", MessageBoxButtons.OK.ToString());		
					}
					catch(Exception ex)
					{
						MessageBox.Show(ex.Message + ": " + command.Parameters["@code"].Value.ToString() + ", " + command.Parameters["@FirstName"].Value.ToString() + " " + command.Parameters["@LastName"].Value.ToString());
					}				
				}
			}
		}

		private void simpleButton4_Click(object sender, System.EventArgs e)
		{
			string sqltxt = @"Insert into ta_ClockTrans ([ClockID], [TrDate], [TrTime], [BadgeNum], [EmpID], [Operation], [RealOperation]) values (@ClockID, @TrDate, @TrTime, @BadgeNum, @EmpID, @Operation, @RealOperation)";
			
			using(SqlConnection connection = new SqlConnection("user id=" + UserID + ";Password=" + Password +";data source=" + ServerName + ";persist security info=True;initial catalog=" + Database))
			{
				using(SqlCommand command = new SqlCommand(sqltxt, connection))
				{	
					try
					{
						command.Parameters.Add("@ClockID" , SqlDbType.Int);
						command.Parameters.Add("@TrDate" , SqlDbType.DateTime);
						command.Parameters.Add("@TrTime", SqlDbType.VarChar);
						command.Parameters.Add("@BadgeNum", SqlDbType.NVarChar);
						command.Parameters.Add("@EmpID", SqlDbType.Int);
						command.Parameters.Add("@Operation", SqlDbType.Char);
						command.Parameters.Add("@RealOperation", SqlDbType.Char);
							
						for (int i = 0; i < gridView3.RowCount; i++)
						{
							//=======================
							//=============EmployeeID
							//=======================
							string sqlEmployeeID = @"SELECT [ID] FROM Emp WHERE code = '" + gridView3.GetRowCellValue(i, gridView3.Columns["emplcode"]).ToString().Trim() + "'";
							SqlCommand cmdEmployeeID = new SqlCommand(sqlEmployeeID, connection);
							connection.Open();
							int empId = Convert.ToInt32(cmdEmployeeID.ExecuteScalar());
							connection.Close();
							//=======================
							//=======================

							//=======================
							//================ClockID
							//=======================
							string sqlClockID = @"SELECT [ID] FROM ta_Clock WHERE code = '" + gridView3.GetRowCellValue(i, gridView3.Columns["clock"]).ToString().Trim() + "'";
							SqlCommand cmdClockID = new SqlCommand(sqlClockID, connection);
							connection.Open();
							int clockId = Convert.ToInt32(cmdClockID.ExecuteScalar());
							connection.Close();
							//=======================
							//=======================

							command.Parameters["@ClockID"].Value = clockId;
							command.Parameters["@TrDate"].Value = gridView3.GetRowCellValue(i,gridView3.Columns["date"]);
							command.Parameters["@TrTime"].Value = gridView3.GetRowCellValue(i,gridView3.Columns["time"]);
							command.Parameters["@BadgeNum"].Value = gridView3.GetRowCellValue(i,gridView3.Columns["emplcode"]);
							command.Parameters["@EmpID"].Value = empId;
							command.Parameters["@Operation"].Value = gridView3.GetRowCellValue(i,gridView3.Columns["function"]);
							command.Parameters["@RealOperation"].Value = gridView3.GetRowCellValue(i,gridView3.Columns["function"]);

							connection.Open();  
							command.ExecuteNonQuery();
							connection.Close();
						}
						MessageBox.Show("It's Done!", MessageBoxButtons.OK.ToString());	
					}
					catch(Exception ex)
					{
						MessageBox.Show(ex.Message);
					}	
				}
			}
		}

		private void simpleButton5_Click(object sender, System.EventArgs e)
		{
			string sqltxt = @"insert into ta_PlannedLeave (EmpID, IsMission, AbsType, OneDayOnly, LeaveDate, Description, WithoutPunch) values (@EmpID, 0, @LeaveID, 1, @LeaveDate, @Desc, 1)";

			using(SqlConnection connection = new SqlConnection("user id=" + UserID + ";Password=" + Password +";data source=" + ServerName + ";persist security info=True;initial catalog=" + Database))
			{
				using(SqlCommand command = new SqlCommand(sqltxt, connection))
				{	
					try
					{
						command.Parameters.Add("@EmpID" , SqlDbType.Int);
						command.Parameters.Add("@LeaveID", SqlDbType.Int);
						command.Parameters.Add("@LeaveDate" , SqlDbType.DateTime);
						command.Parameters.Add("@Desc", SqlDbType.NVarChar);
							
						for (int i = 0; i < gridView3.RowCount; i++)
						{
							//=======================
							//=============EmployeeID
							//=======================
							string sqlEmployeeID = @"SELECT [ID] FROM Emp WHERE code = '" + gridView3.GetRowCellValue(i, gridView3.Columns["emplcode"]).ToString().Trim() + "'";
							SqlCommand cmdEmployeeID = new SqlCommand(sqlEmployeeID, connection);
							connection.Open();
							int empId = Convert.ToInt32(cmdEmployeeID.ExecuteScalar());
							connection.Close();
							//=======================
							//=======================

							//=======================
							//================LeaveID
							//=======================
							string sqlLeaveID = @"SELECT [ID] FROM Absence where Abbreviation = '" + gridView3.GetRowCellValue(i, gridView3.Columns["category"]).ToString().Trim() + "'";
							SqlCommand cmdLeaveID = new SqlCommand(sqlLeaveID, connection);
							connection.Open();
							int leaveId = Convert.ToInt32(cmdLeaveID.ExecuteScalar());
							connection.Close();
							//=======================
							//=======================

							//=======================
							//==============LeaveDesc
							//=======================
							string sqlLeaveDesc = @"SELECT [Description] FROM Absence where Abbreviation = '" + gridView3.GetRowCellValue(i, gridView3.Columns["category"]).ToString().Trim() + "'";
							SqlCommand cmdLeaveDesc = new SqlCommand(sqlLeaveDesc, connection);
							connection.Open();
							string leaveDesc = Convert.ToString(cmdLeaveDesc.ExecuteScalar());
							connection.Close();
							//=======================
							//=======================

							command.Parameters["@EmpID"].Value = empId;
							command.Parameters["@LeaveID"].Value = leaveId;
							command.Parameters["@LeaveDate"].Value = gridView3.GetRowCellValue(i,gridView3.Columns["date"]);
							command.Parameters["@Desc"].Value = gridView3.GetRowCellValue(i, gridView3.Columns["emplcode"]).ToString().Trim() + " - " + leaveDesc;

							connection.Open();  
							command.ExecuteNonQuery();
							connection.Close();
							
							//=======================
							//==============LeaveType
							//=======================
							string sqlLeaveType = @"SELECT [ded_Leave_Type] FROM Absence where Abbreviation = '" + gridView3.GetRowCellValue(i, gridView3.Columns["category"]).ToString().Trim() + "'";
							SqlCommand cmdLeaveType = new SqlCommand(sqlLeaveType, connection);
							connection.Open();
							string leaveType = Convert.ToString(cmdLeaveType.ExecuteScalar());
							connection.Close();
							//=======================
							//=======================

							//=======================
							//====ta_PlannedLeaveDays
							//=======================
							if(leaveType == "Y")
							{
								string sqlPLID = @"SELECT max(ID) FROM ta_PlannedLeave";
								SqlCommand cmdPLID = new SqlCommand(sqlPLID, connection);
								connection.Open();
								int PLID = Convert.ToInt32(cmdPLID.ExecuteScalar());
								connection.Close();

								string sqlLeaveDays = @"insert into ta_PlannedLeaveDays (PLID, PLDate, PLDOW, NbDaysOrig, NbDaysEdit) values (@PLID, @LeaveDate, CASE DATEPART(WEEKDAY, @LeaveDate) WHEN 1 THEN 'Sunday' WHEN 2 THEN 'Monday' WHEN 3 THEN 'Tuesday' WHEN 4 THEN 'Wednesday' WHEN 5 THEN 'Thursday' WHEN 6 THEN 'Friday' ELSE 'Saturday' END, 1, 1)";
								SqlCommand cmdLeaveDays = new SqlCommand(sqlLeaveDays, connection);

								cmdLeaveDays.Parameters.Add("@PLID", SqlDbType.Int);
								cmdLeaveDays.Parameters.Add("@LeaveDate" , SqlDbType.DateTime);

								cmdLeaveDays.Parameters["@PLID"].Value = PLID;
								cmdLeaveDays.Parameters["@LeaveDate"].Value = gridView3.GetRowCellValue(i,gridView3.Columns["date"]);

								connection.Open();
								cmdLeaveDays.ExecuteNonQuery();
								connection.Close();
							}
							//=======================
							//=======================
						}
						MessageBox.Show("It's Done!", MessageBoxButtons.OK.ToString());	
					}
					catch(Exception ex)
					{
						MessageBox.Show(ex.Message);
					}	
				}
			}
		}
	}
}


