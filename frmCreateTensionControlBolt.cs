/*--------------------------------------------------------------------------------------+
|
|     GSF - Garrison Steel Fabricators, Pell city, AL
|     Developer - Albert Sharapov, Steel Detailer
|
|     Bentley Prosteel Addin "GSF Tension Control Bolt"
|     creates bolts with sphere heads, 
|     converts standard bolts to bolts with shpere heads
|
+--------------------------------------------------------------------------------------*/

using System;
using System.IO;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Generic;
using SRI=System.Runtime.InteropServices;

using Bentley.ProStructures;
using ProStructuresOE;
using Bentley.ProStructures.CadSystem;
using Bentley.ProStructures.Configuration;
using Bentley.ProStructures.Drawing;
using Bentley.ProStructures.Geometry.Data;
using Bentley.ProStructures.Geometry.Utilities;
using Bentley.ProStructures.Miscellaneous;
using Bentley.ProStructures.Property;
using Bentley.ProStructures.Entity;
using Bentley.ProStructures.Steel.Plate;
using Bentley.ProStructures.Steel.Primitive;
using Bentley.ProStructures.Steel.Bolt;
using Bentley.ProStructures.Modification.Edit;
using Bentley.ProStructures.Modification.ObjectData;
using Microsoft.VisualBasic;
using PlugInBase;

using BM=Bentley.MicroStation;
using BMW=Bentley.MicroStation.WinForms;
using BMI=Bentley.MicroStation.InteropServices;
using BCOM=Bentley.Interop.MicroStationDGN;

namespace GSFBolt
{
    public class BoltControl : BMW.Adapter 
    {
        private static BoltControl              s_current;
        private Bentley.Windowing.WindowContent  m_windowContent;

        private Bentley.MicroStation.AddIn                        m_addIn;

        private System.Windows.Forms.GroupBox 			boltSettings;
        private System.Windows.Forms.GroupBox           anchorSettings;
        private System.Windows.Forms.TabPage            tcBoltSettings;
        private System.Windows.Forms.TabPage            expAnchorSettings;
        private System.Windows.Forms.TabControl         tabControl1;
        private System.Windows.Forms.Label 				boltTypes;
        private System.Windows.Forms.Label 				boltDiameters;
        private System.Windows.Forms.Label 				boltAdditionalLength;
        private System.Windows.Forms.ComboBox 			boltTypesComboBox;
        private System.Windows.Forms.ComboBox 			boltDiametersComboBox;
        private System.Windows.Forms.TextBox 			boltAdditionalLengthTextBox;
        private System.Windows.Forms.Label              anchorNameLabel;
        private System.Windows.Forms.Label 				anchorDiameters;
        private System.Windows.Forms.Label              anchorEmbedmentLabel;
        private System.Windows.Forms.Label              partThicknessLabel;
        private System.Windows.Forms.Label              anchorLengthLabel;
        private System.Windows.Forms.ComboBox 			anchorNameComboBox;
        private System.Windows.Forms.ComboBox 			anchorDiametersComboBox;
        private System.Windows.Forms.TextBox 			anchorEmbedmentTextBox;
        private System.Windows.Forms.TextBox 			partThicknessTextBox;
        private System.Windows.Forms.ComboBox 			anchorLengthComboBox;
        private System.Windows.Forms.Button 		   	insertSingleBolt;
        private System.Windows.Forms.Button   			cancelButton;
        private System.Windows.Forms.Button   			okButton;
        private System.Windows.Forms.Button 			convertBolt;
        private System.Windows.Forms.Button 			helpBolt;
        private System.Windows.Forms.Button 			insertExpansionAnchor;
        private System.Windows.Forms.Button 			anchorDbConnection;
        private System.Windows.Forms.Button 			helpAnchor;
        private System.Windows.Forms.Button 			cancelButton1;
        private System.Windows.Forms.Button   			okButton1;
        private System.Windows.Forms.ToolTip 			toolTip1;
        private System.Windows.Forms.CheckBox           anchorExpansionCheckBox;
        private System.Windows.Forms.CheckBox           anchorLengthCheckBox;
        private System.Windows.Forms.Label              anchorEmbedmentValuesRange;

        private static int                              tabPage;
        private static int                              objId;
        private static string[]                         anchorboltdiameters_imperial;
        private static string[]                         anchorboltlengths_imperial;
        private static bool 							anchorEmbedmentValueVerified;
        private System.Windows.Forms.Form               form1;
        private	System.Windows.Forms.LinkLabel 			websitelinklabel;
        private	System.Windows.Forms.LinkLabel 			techdatalinklabel;

        /// <summary>The Visual Studio IDE requires a constructor without arguments.</summary>
        private BoltControl()
        {
            System.Diagnostics.Debug.Assert (this.DesignMode, "Do not use the default constructor");
            InitializeComponent(tabPage);
        }
            
        /// <summary>Constructor</summary>
        internal BoltControl(Bentley.MicroStation.AddIn addIn)
        {
            m_addIn     = addIn;
            InitializeComponent(tabPage);

            //  Set up events to handle closing of form
            this.Closed += new EventHandler(BoltControl_Closed);
        }

        /// Clean up any resources being used.
        protected override void Dispose( bool disposing )
        {
            base.Dispose( disposing );
        }

        private void InitializeComponent(int TabIndex)
        {
            this.boltSettings = new System.Windows.Forms.GroupBox();
            this.anchorSettings = new System.Windows.Forms.GroupBox();
            this.tcBoltSettings = new System.Windows.Forms.TabPage();
            this.expAnchorSettings = new System.Windows.Forms.TabPage();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.boltDiametersComboBox = new System.Windows.Forms.ComboBox();
            this.boltTypesComboBox = new System.Windows.Forms.ComboBox();
            this.boltAdditionalLengthTextBox = new System.Windows.Forms.TextBox();
            this.boltDiameters = new System.Windows.Forms.Label();
            this.boltTypes = new System.Windows.Forms.Label();
            this.boltAdditionalLength = new System.Windows.Forms.Label();
            this.anchorNameLabel = new System.Windows.Forms.Label();
            this.anchorDiameters = new System.Windows.Forms.Label();
            this.anchorEmbedmentLabel = new System.Windows.Forms.Label();
            this.partThicknessLabel = new System.Windows.Forms.Label();
            this.anchorLengthLabel = new System.Windows.Forms.Label();
            this.anchorNameComboBox = new System.Windows.Forms.ComboBox();
            this.anchorDiametersComboBox = new System.Windows.Forms.ComboBox();
            this.anchorEmbedmentTextBox = new System.Windows.Forms.TextBox();
            this.partThicknessTextBox = new System.Windows.Forms.TextBox();
            this.anchorLengthComboBox = new System.Windows.Forms.ComboBox();
            this.insertSingleBolt = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.convertBolt = new System.Windows.Forms.Button();
            this.helpBolt = new System.Windows.Forms.Button();
            this.insertExpansionAnchor = new System.Windows.Forms.Button();
            this.anchorDbConnection = new System.Windows.Forms.Button();
            this.helpAnchor = new System.Windows.Forms.Button();
            this.cancelButton1 = new System.Windows.Forms.Button();
            this.okButton1 = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip();
            this.anchorExpansionCheckBox = new System.Windows.Forms.CheckBox();
            this.anchorLengthCheckBox = new System.Windows.Forms.CheckBox();
            this.anchorEmbedmentValuesRange = new System.Windows.Forms.Label();
            this.boltSettings.SuspendLayout();
            this.SuspendLayout();

            //
            // Properties
            //
            //
            //Buttons size and location
            //
            System.Drawing.Size buttonSize = new System.Drawing.Size (33, 33);
            int space_bw_buttons = 5;
            int space_bw_buttons_and_groupbox = 8;
            int step_buttonseries = buttonSize.Width + space_bw_buttons;
            //
            //Label size and location
            //
            System.Drawing.Size labelSize = new System.Drawing.Size(100, 20);
            System.Drawing.Point firstPointlabelseries = new System.Drawing.Point(5, 20);
            int space_bw_labels = 5;
            int step_labelseries = labelSize.Height + space_bw_labels;
            System.Drawing.Size embedmentValuesLabelSize = new System.Drawing.Size(80, 40);
            //
            //Combobox size and location
            //
            System.Drawing.Size comboboxSize = new System.Drawing.Size (130, 20);
            int space_bw_label_and_combobox = 1;
            System.Drawing.Point firstPointcomboboxseries = new System.Drawing.Point(
                firstPointlabelseries.X+ labelSize.Width + space_bw_label_and_combobox,
                firstPointlabelseries.Y);
            //
            //TextBox Size
            //
            System.Drawing.Size textboxSize = new System.Drawing.Size(100, 20);
            //
            //CheckBoxes
            //
            System.Drawing.Point firstPointcheckboxseries = new System.Drawing.Point(
            	firstPointlabelseries.X + labelSize.Width + space_bw_label_and_combobox + textboxSize.Width + 5,
            	firstPointlabelseries.Y + 2*labelSize.Height + 2*space_bw_labels);
            System.Drawing.Size checkboxSize = new System.Drawing.Size(20, 20);
            //
            //Locatition of first button
            //
            System.Drawing.Point firstPointbuttonseries = new System.Drawing.Point(
                12,
                firstPointlabelseries.Y + 5*labelSize.Height + 5*space_bw_labels + space_bw_buttons_and_groupbox);
            //
            // Load values
            //
            string[] types = new string[]{"SHOP TC A325","FIELD TC A325", "SHOP TC A490", "FIELD TC A490"};
            this.boltTypesComboBox.Items.AddRange(types);
            string[] tcboltdiameters_imperial = new string[]
                {
                    "0:0 1/2","0:0 5/8", "0:0 3/4", "0:0 7/8", "0:1", "0:1 1/8"
                };
            this.boltDiametersComboBox.Items.AddRange(tcboltdiameters_imperial);
            string[] arr = BoltDb.TablesToArray();
            this.anchorNameComboBox.Items.AddRange(arr);
            
            PsUnits psUnits = new PsUnits();
            this.anchorDiametersComboBox.Items.Clear();
            this.anchorLengthComboBox.Items.Clear();
            PsTemplateManager psTemplateManager = new PsTemplateManager();
            psTemplateManager.LoadTemplateEntry("GSFBolt");
            if (psTemplateManager.IsLoaded)
                {
                long num1 = (long)psTemplateManager.get_Number(0);
                this.boltTypesComboBox.Text = psTemplateManager.get_String(0);
                this.boltDiametersComboBox.Text = psTemplateManager.get_String(1);
                this.boltAdditionalLengthTextBox.Text = psTemplateManager.get_String(2);
                this.anchorNameComboBox.Text = psTemplateManager.get_String(3);
                this.anchorDiametersComboBox.Text = psTemplateManager.get_String(4);
                this.anchorEmbedmentTextBox.Text = psTemplateManager.get_String(5);
                this.partThicknessTextBox.Text = psTemplateManager.get_String(6);
                this.anchorLengthComboBox.Text = psTemplateManager.get_String(7);
                this.anchorLengthCheckBox.Checked = psTemplateManager.get_Boolean(0);
                int pos = Array.IndexOf(arr, this.anchorNameComboBox.Text.ToString());
        		if (pos < 1)
        			{
        				anchorboltdiameters_imperial = new string[]
        		    	{
        		    		"0:0 1/4","0:0 3/8","0:0 1/2","0:0 5/8", "0:0 3/4", "0:0 7/8", "0:1", "0:1 1/8", "0:1 1/4"
        		    	};
        				this.anchorDiametersComboBox.Items.AddRange(anchorboltdiameters_imperial);
        				this.anchorDiametersComboBox.Text = psTemplateManager.get_String(4);
        				this.anchorLengthComboBox.Text = psTemplateManager.get_String(7);
        			}
        		else
        			{
        			anchorboltdiameters_imperial = BoltDb.PopulateComboBoxWithDiameters(
        				this.anchorNameComboBox.Text.ToString()
        				);
        			this.anchorDiametersComboBox.Items.AddRange(anchorboltdiameters_imperial);
        			this.anchorDiametersComboBox.Text = psTemplateManager.get_String(4);
        			double criteria = psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text);
        			anchorboltlengths_imperial = BoltDb.PopulateComboBoxWithLengths(
        				this.anchorNameComboBox.Text.ToString(),
        				criteria);
        			this.anchorLengthComboBox.Items.AddRange(anchorboltlengths_imperial);
        			this.anchorLengthComboBox.Text = psTemplateManager.get_String(7);
        			}
                }
            else
                {
                	anchorboltdiameters_imperial = new string[]
        		    	{
        		    		"0:0 1/4","0:0 3/8","0:0 1/2","0:0 5/8", "0:0 3/4", "0:0 7/8", "0:1", "0:1 1/8", "0:1 1/4"
        		    	};
        			this.anchorDiametersComboBox.Items.AddRange(anchorboltdiameters_imperial);
                    this.boltTypesComboBox.Text = types[0];
                    this.boltDiametersComboBox.Text = tcboltdiameters_imperial[0];
                    this.boltAdditionalLengthTextBox.Text = "0:0";
                    this.anchorNameComboBox.Text = arr[0];
                    this.anchorDiametersComboBox.Text = anchorboltdiameters_imperial[0];
                    this.anchorEmbedmentTextBox.Text = "0:0";
                    this.partThicknessTextBox.Text = "0:0";
                    this.anchorLengthComboBox.Text = "0:0";
                    this.anchorLengthCheckBox.Checked = true;
                }    
        	//
            // Create the ToolTip
            //
            this.toolTip1.AutoPopDelay = 5000;
            this.toolTip1.InitialDelay = 1000;
            this.toolTip1.ReshowDelay = 500;
            // 
            // boltSettings
            //
            this.boltSettings.Text = "Bolt Settings";
            this.boltSettings.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left |
                System.Windows.Forms.AnchorStyles.Top)));
            this.boltSettings.Location = new System.Drawing.Point (8, 5);
            this.boltSettings.Size = new System.Drawing.Size (
                2*firstPointlabelseries.X + labelSize.Width + space_bw_label_and_combobox + comboboxSize.Width,
                firstPointlabelseries.Y + 3*labelSize.Height + 3*space_bw_labels);
            // this.boltSettings.TabIndex = 0;
            //
            // anchorSettings
            //
            this.anchorSettings.Text = "Anchor Settings";
            this.anchorSettings.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left |
                System.Windows.Forms.AnchorStyles.Top)));
            this.anchorSettings.Location = new System.Drawing.Point (8, 5);
            this.anchorSettings.Size = new System.Drawing.Size (
            	2*firstPointlabelseries.X + labelSize.Width + space_bw_label_and_combobox + comboboxSize.Width,
                firstPointlabelseries.Y + 5*labelSize.Height + 5*space_bw_labels
                );
            //
            //tcBoltSettings
            //
            this.tcBoltSettings.Text = "TC Bolt";
            this.tcBoltSettings.TabIndex = 0;
            //
            //anchorSettings
            //
            this.expAnchorSettings.Text = "Expansion Anchor";
            this.expAnchorSettings.TabIndex = 1;
            //
            //tabControl1
            //
            this.tabControl1.Location = new System.Drawing.Point(6, 7);
            this.tabControl1.Size = new System.Drawing.Size(
            	2 + 3*this.boltSettings.Location.X + this.boltSettings.Size.Width,
            	23 + firstPointbuttonseries.Y + buttonSize.Height + 5
            	);
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.TabIndex = 0;  
            //
            //boltTypes label
            //
            this.boltTypes.Text = "Bolt type";
            this.boltTypes.Size = labelSize;
            this.boltTypes.Location = firstPointlabelseries;
            this.boltTypes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //
            //boltDiameters label
            //
            this.boltDiameters.Text = "Diameter";
            this.boltDiameters.Size = labelSize;
            this.boltDiameters.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + step_labelseries);
            this.boltDiameters.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //boltAdditionalLength label
            //
            this.boltAdditionalLength.Text = "Additional Length";
            this.boltAdditionalLength.Size = labelSize;
            this.boltAdditionalLength.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + 2*step_labelseries);
            this.boltAdditionalLength.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //boltTypesComboBox
            //
            
            this.boltTypesComboBox.Size = comboboxSize;
            this.boltTypesComboBox.Location = firstPointcomboboxseries;
            this.boltTypesComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //
            //boltDiameterComboBox
            //
            this.boltDiametersComboBox.Size = comboboxSize;
            this.boltDiametersComboBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + step_labelseries);
            this.boltDiametersComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //
            // boltAdditionalLengthTextBox
            //
            this.boltAdditionalLengthTextBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + 2*step_labelseries);
            //
            //Anchor label
            //
            this.anchorNameLabel.Text = "Anchor";
            this.anchorNameLabel.Size = labelSize;
            this.anchorNameLabel.Location = firstPointlabelseries;
            this.anchorNameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //
            //anchorDiameters label
            //
            this.anchorDiameters.Text = "Diameter";
            this.anchorDiameters.Size = labelSize;
            this.anchorDiameters.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + step_labelseries);
            this.anchorDiameters.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //Anchor Embedment label
            //

            this.anchorEmbedmentLabel.Text = "Embedment";
            this.anchorEmbedmentLabel.Size = labelSize;
            this.anchorEmbedmentLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + 2*step_labelseries);
            this.anchorEmbedmentLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //Anchor Length label
            //
            this.partThicknessLabel.Text = "Part Thickness";
            this.partThicknessLabel.Size = labelSize;
            this.partThicknessLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + 3*step_labelseries);
            this.partThicknessLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //Anchor Length label
            //
            this.anchorLengthLabel.Text = "Length";
            this.anchorLengthLabel.Size = labelSize;
            this.anchorLengthLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + 4*step_labelseries);
            this.anchorLengthLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //anchorNameComboBox
            //
            this.anchorNameComboBox.Size = comboboxSize;
            this.anchorNameComboBox.Location = firstPointcomboboxseries;
            this.anchorNameComboBox.SelectedIndexChanged += new EventHandler(this.AnchorNameComboBox_SelectedIndexChanged);
            this.anchorNameComboBox.TextChanged += new EventHandler(this.AnchorNameComboBox_SelectedIndexChanged);
            // this.anchorNameComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //
            //anchorDiametersComboBox
            //
            this.anchorDiametersComboBox.Size = comboboxSize;
            this.anchorDiametersComboBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + step_labelseries);
            this.anchorDiametersComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.anchorDiametersComboBox.SelectedIndexChanged += new EventHandler(this.AnchorDiameterComboBox_SelectedIndexChanged);
            this.anchorDiametersComboBox.SelectedIndexChanged += new EventHandler(this.AnchorEmbedmentValueVerify);
            //
            // anchorEmbedmentTextBox
            //
            this.anchorEmbedmentTextBox.Width = 50;
            this.anchorEmbedmentTextBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + 2*step_labelseries);
            this.anchorEmbedmentTextBox.TextChanged += new EventHandler(this.AnchorEmbedmentValueVerify);
            // this.anchorEmbedmentTextBox.Enabled = this.anchorEmbedmentCheckBox.Checked;
            //
            // partThicknessTextBox
            //
            this.partThicknessTextBox.Width = 50;
            this.partThicknessTextBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + 3*step_labelseries);
            this.partThicknessTextBox.TextChanged += new EventHandler(this.AnchorEmbedmentValueVerify);
            //
            // anchorLengthComboBox
            //
            this.anchorLengthComboBox.Width = 100;
            this.anchorLengthComboBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + 4*step_labelseries);
            this.anchorLengthComboBox.Enabled = this.anchorLengthCheckBox.Checked;
            this.anchorLengthComboBox.SelectedIndexChanged += new EventHandler(this.AnchorEmbedmentValueVerify);
            this.anchorLengthComboBox.TextChanged += new EventHandler(this.AnchorEmbedmentValueVerify);
            //
            //anchorLengthCheckBox
            //
            this.anchorLengthCheckBox.Location = new System.Drawing.Point(
            	firstPointcheckboxseries.X,
            	firstPointcheckboxseries.Y + 2*step_labelseries
            	);
            this.anchorLengthCheckBox.Size = checkboxSize;
            this.anchorLengthCheckBox.CheckStateChanged += new System.EventHandler(this.LengthCheckBox_CheckStateChanged);
            this.anchorLengthCheckBox.CheckStateChanged += new EventHandler(this.AnchorEmbedmentValueVerify);
            //
            //Anchor Length label
            //
            this.anchorEmbedmentValuesRange.Size = embedmentValuesLabelSize;
            this.anchorEmbedmentValuesRange.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X + this.anchorEmbedmentTextBox.Width + 2,
                firstPointlabelseries.Y + 2*step_labelseries + 2);
            this.anchorEmbedmentValuesRange.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            anchorEmbedmentValueVerified = anchorEmbedmentValuesVerified();
            //
            // Load Bitmap Images
            //
            System.Reflection.Assembly ass1 = System.Reflection.Assembly.GetExecutingAssembly();
            Stream stream1 = ass1.GetManifestResourceStream("GSFBolt.tcBoltIcon.bmp");
            Stream stream2 = ass1.GetManifestResourceStream("GSFBolt.ToTCBoltIcon.ico");
            Stream stream3 = ass1.GetManifestResourceStream("GSFBolt.InfoIcon.ico");
            Stream stream4 = ass1.GetManifestResourceStream("GSFBolt.CancelIcon.ico");
            Stream stream5 = ass1.GetManifestResourceStream("GSFBolt.OkIcon.ico");
            // 
            // insertSingleBolt
            //
            this.insertSingleBolt.Name = "insertSingleBolt";
            this.insertSingleBolt.Image = Image.FromStream(stream1);
            this.toolTip1.SetToolTip(this.insertSingleBolt, "Create Single Tension Control Bolt By Two Points");
            this.insertSingleBolt.Size = buttonSize;
            this.insertSingleBolt.Location = firstPointbuttonseries;
            this.insertSingleBolt.Click += new System.EventHandler(this.cmdInsertSingleBolt);
            //
            // convertBolt
            //
            this.convertBolt.Name = "convertBolt";
            this.convertBolt.Image = Image.FromStream(stream2);
            this.toolTip1.SetToolTip(this.convertBolt, "Convert to TC Bolt");
            this.convertBolt.Size = buttonSize;
            this.convertBolt.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + step_buttonseries,
                firstPointbuttonseries.Y);
            this.convertBolt.Click += new System.EventHandler(this.cmdConvertBolt);
            //
            // OkButton
            //
            this.okButton.Name = "Ok";
            this.okButton.Image = Image.FromStream(stream5);
            this.toolTip1.SetToolTip(this.okButton, "OK");
            this.okButton.Size = buttonSize;
            this.okButton.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 3*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.okButton.Click += new System.EventHandler(this.cmdOk);
            //
            // CancelButton
            //
            this.cancelButton.Name = "Cancel";
            this.cancelButton.Image = Image.FromStream(stream4);
            this.toolTip1.SetToolTip(this.cancelButton, "Cancel");
            this.cancelButton.Size = buttonSize;
            this.cancelButton.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 4*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.cancelButton.Click += new System.EventHandler(this.cmdCancel);
            //
            //helpBolt
            //
            this.helpBolt.Name = "helpBolt";
            this.helpBolt.Image = Image.FromStream(stream3);
            this.toolTip1.SetToolTip(this.helpBolt, "About");
            this.helpBolt.Size = buttonSize;
            this.helpBolt.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 5*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.helpBolt.Click += new System.EventHandler(this.cmdHelpBolt);
            // 
            // insertSingleAnchor
            //
            this.insertExpansionAnchor.Name = "insertExpansionAnchor";
            this.insertExpansionAnchor.Image = Image.FromStream(stream1);
            this.toolTip1.SetToolTip(this.insertExpansionAnchor, "Create Single Expansion Anchor" + Environment.NewLine + 
            	"using Insert Point and Direction" + Environment.NewLine +
            	"or using Insert Point and Part Thickness");
            this.insertExpansionAnchor.Size = buttonSize;
            this.insertExpansionAnchor.Location = firstPointbuttonseries;
            this.insertExpansionAnchor.Click += new System.EventHandler(this.cmdInsertAnchor);
            // 
            // anchorDbConnection
            //
            this.anchorDbConnection.Name = "anchorDbConnection";
            this.anchorDbConnection.Size = buttonSize;
            this.anchorDbConnection.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + step_buttonseries,
                firstPointbuttonseries.Y);
            this.anchorDbConnection.Click += new System.EventHandler(this.cmdAnchorData);
            //
            // OkButton
            //
            this.okButton1.Name = "Ok";
            this.okButton1.Image = Image.FromStream(stream5);
            this.toolTip1.SetToolTip(this.okButton1, "OK");
            this.okButton1.Size = buttonSize;
            this.okButton1.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 3*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.okButton1.Click += new System.EventHandler(this.cmdOk);
            //
            // CancelButton1
            //
            this.cancelButton1.Name = "Cancel";
            this.cancelButton1.Image = Image.FromStream(stream4);
            this.toolTip1.SetToolTip(this.cancelButton1, "Cancel");
            this.cancelButton1.Size = buttonSize;
            this.cancelButton1.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 4*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.cancelButton1.Click += new System.EventHandler(this.cmdCancel);
            //
            //helpBolt
            //
            this.helpAnchor.Name = "helpAnchor";
            this.helpAnchor.Image = Image.FromStream(stream3);
            this.toolTip1.SetToolTip(this.helpAnchor, "About");
            this.helpAnchor.Size = buttonSize;
            this.helpAnchor.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 5*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.helpAnchor.Click += new System.EventHandler(this.cmdHelpBolt);
            //
            // Groupbox boltSettings Controls
            //
            this.boltSettings.Controls.Add(this.boltTypes);
            this.boltSettings.Controls.Add(this.boltDiameters);
            this.boltSettings.Controls.Add(this.boltAdditionalLength);
            this.boltSettings.Controls.Add(this.boltTypesComboBox);
            this.boltSettings.Controls.Add(this.boltDiametersComboBox);
            this.boltSettings.Controls.Add(this.boltAdditionalLengthTextBox);
            //
            // Groupbox anchorSettings Controls
            //
            this.anchorSettings.Controls.Add(this.anchorNameLabel);
            this.anchorSettings.Controls.Add(this.anchorDiameters);
            this.anchorSettings.Controls.Add(this.anchorEmbedmentLabel);
            this.anchorSettings.Controls.Add(this.partThicknessLabel);
            this.anchorSettings.Controls.Add(this.anchorLengthLabel);
            this.anchorSettings.Controls.Add(this.anchorNameComboBox);
            this.anchorSettings.Controls.Add(this.anchorDiametersComboBox);
            this.anchorSettings.Controls.Add(this.anchorEmbedmentTextBox);
            this.anchorSettings.Controls.Add(this.partThicknessTextBox);
            this.anchorSettings.Controls.Add(this.anchorLengthComboBox);
            this.anchorSettings.Controls.Add(this.anchorLengthCheckBox);
            this.anchorSettings.Controls.Add(this.anchorEmbedmentValuesRange);
            //
            // TabPage tcBoltSettings Controls
            //
            this.tcBoltSettings.Controls.Add(this.boltSettings);
            this.tcBoltSettings.Controls.Add(this.insertSingleBolt);
            this.tcBoltSettings.Controls.Add(this.okButton);
            this.tcBoltSettings.Controls.Add(this.cancelButton);
            this.tcBoltSettings.Controls.Add(this.convertBolt);
            this.tcBoltSettings.Controls.Add(this.helpBolt);
            //
            // TabPage expAnchorSettings Controls
            //
            this.expAnchorSettings.Controls.Add(this.anchorSettings);
            this.expAnchorSettings.Controls.Add(this.insertExpansionAnchor);
            this.expAnchorSettings.Controls.Add(this.anchorDbConnection);
            this.expAnchorSettings.Controls.Add(this.okButton1);
            this.expAnchorSettings.Controls.Add(this.cancelButton1);
            this.expAnchorSettings.Controls.Add(this.helpAnchor);
            //
            //TabControl
            //
            this.tabControl1.Controls.Add(this.tcBoltSettings);
            this.tabControl1.Controls.Add(this.expAnchorSettings);
            //keeps tabindex after closing form
            this.tabControl1.SelectedIndex = TabIndex;
            //
            //Add control to the form
            //
            this.Controls.Add(this.tabControl1);
            this.Name = "ProSteel AddIn";
            this.Text = "GSF TC Bolts/Expansion Anchors";
            this.boltSettings.ResumeLayout(false);
            this.ResumeLayout(false);
        }
            
        #endregion
        /// <summary>
        /// Show the form if it is not already displayed
        /// </summary>
        internal static void ShowForm (Bentley.MicroStation.AddIn addIn)
        {
            if (null != s_current)
                return;

            s_current = new BoltControl (addIn);
            s_current.AttachAsTopLevelForm (addIn, true);

            
            s_current.AutoOpen = true;
            s_current.AutoOpenKeyin = "mdl load GSFBolt";
            
            s_current.NETDockable = false; 
            Bentley.Windowing.WindowManager    windowManager = 
                        Bentley.Windowing.WindowManager.GetForMicroStation ();
            s_current.m_windowContent = 
                windowManager.DockPanel (s_current, s_current.Name, s_current.Text, 
                Bentley.Windowing.DockLocation.Floating);
            AdapterWorkarounds.WCFixedBorder.SetBorderFixed(s_current.m_windowContent);

            s_current.m_windowContent.CanDockHorizontally = false; // limit to left and right docking
            s_current.m_windowContent.CanDockVertically = false;
        }

        /// <summary>
        /// Close the form if it is currently displayed
        /// </summary>
        internal static void CloseForm ()
        {
            if (s_current == null)
                return;
            s_current.m_windowContent.Close();

            s_current = null;
        }


        /// <summary>Handle the standard Closed event
        /// </summary>
        private void BoltControl_Closed(object sender, EventArgs e)
       {   
           if (s_current != null)
           tabPage = this.tabControl1.SelectedIndex;
               s_current.Dispose (true);
              
           s_current = null;
       }

        private void cmdInsertSingleBolt(object sender, System.EventArgs e)
        {
            PsUnits psUnits = new PsUnits();
            string boltType = this.boltTypesComboBox.SelectedItem.ToString();
            double boltDiameter = psUnits.ConvertToNumeric(this.boltDiametersComboBox.Text);
            double boltAddLength = psUnits.ConvertToNumeric(this.boltAdditionalLengthTextBox.Text);
    		PsTemplateManager psTemplateManager = new PsTemplateManager();
            psTemplateManager.AppendNumber(1);
            psTemplateManager.AppendString(this.boltTypesComboBox.SelectedItem.ToString());
            psTemplateManager.AppendString(this.boltDiametersComboBox.SelectedItem.ToString());
            psTemplateManager.AppendString(this.boltAdditionalLengthTextBox.Text);
            psTemplateManager.AppendString(this.anchorNameComboBox.Text.ToString());
            psTemplateManager.AppendString(this.anchorDiametersComboBox.SelectedItem.ToString());
            psTemplateManager.AppendString(this.anchorEmbedmentTextBox.Text);
            psTemplateManager.AppendString(this.partThicknessTextBox.Text);
            psTemplateManager.AppendString(this.anchorLengthComboBox.Text.ToString());
            psTemplateManager.AppendBoolean(this.anchorLengthCheckBox.Checked);
            psTemplateManager.SaveTemplateEntry("GSFBolt", 1);
    		BoltControl.CloseForm();
            this.CreateSingleBolt(boltType, boltDiameter, boltAddLength);
            BoltControl.ShowForm(m_addIn);
        }

        private void CreateSingleBolt(string BoltType, double BoltDiameter, double BoltAddLength)
        {
        	TcBolt tcBolt = new TcBolt();
        	tcBolt.SetParams(BoltDiameter, BoltType, BoltAddLength);

            PsPoint psStartPoint = new PsPoint();
            PsPoint psEndPoint = new PsPoint();
            PsPoint reference = new PsPoint();
            
            PsPoint psPoint = new PsPoint();
            PsMatrix psMatrix = new PsMatrix();
            PsSelection psSelection = new PsSelection();
            
            PsSelection arg1 = psSelection;
            int num1 = arg1.PickPoint("Select start point...", CoordSystem.kWcs, psStartPoint);
            if (num1 == 1)
            {
                reference = psMatrix.TransformPoint(psStartPoint);
                PsSelection arg2 = psSelection;
                num1 = arg2.PickPointWithReference("Select end point...", CoordSystem.kWcs, reference, psEndPoint);
                if (num1 == 1)
                {
                	tcBolt.CreateByTwoPoints(psStartPoint, psEndPoint);
                }
            }
        }

        private void cmdConvertBolt(object sender, System.EventArgs e)
    	{
    		BoltControl.CloseForm();
    		Convert();
    		BoltControl.ShowForm(m_addIn);
    	}

        private void Convert()
    	{
    		PsSelection arg1 = new PsSelection();
    		arg1.SetSelectionFilter(SelectionFilter.kFilterBolts);
    		int num1 = arg1.SelectObjects("Select bolts for replacement...");
            if (num1 >0)
            {
            	PsObjectProperties psObjectProperties = new PsObjectProperties();
    			int objectCount = arg1.ObjectCount;
    			for (int j = 0; j <= objectCount - 1; j++)
    			{
    				int b_id = arg1.get_Object(j);
            		psObjectProperties.readFrom(b_id);
            		if (psObjectProperties.ObjectType == ObjectType.kBolt)
            		{
            			TcBolt tcBolt = new TcBolt();
            			tcBolt.ConvertToTCBolt(b_id);
            		}
            		else if (1 == 1)
            		{
    					string message = 
    					"Cannot convert Object with Id: " + num1 + " to TC Bolt. Try to choose another object..." + Environment.NewLine +
    					Environment.NewLine;
    					string caption = "GSF TC Bolt/Exp. Anchor";
    					MessageBoxButtons buttons = MessageBoxButtons.OK;
    					DialogResult result;
    					result = MessageBox.Show(message, caption, buttons); 
            		}
            	}	
            }
    	}
        private void cmdInsertAnchor(object sender, System.EventArgs e)
        {
        	if (this.anchorLengthCheckBox.Checked)
        	{
    	    	if (anchorEmbedmentValueVerified)
    	    	{
    	        	PsUnits psUnits = new PsUnits();
    	        	string anchorName = this.anchorNameComboBox.Text;
    	        	double anchorDiameter = psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text);
    	        	double embedment = psUnits.ConvertToNumeric(this.anchorEmbedmentTextBox.Text);
    	        	double partThickness = psUnits.ConvertToNumeric(this.partThicknessTextBox.Text);
    	        	double anchorLength = psUnits.ConvertToNumeric(this.anchorLengthComboBox.Text);
    	        	PsTemplateManager psTemplateManager = new PsTemplateManager();
    	        	psTemplateManager.AppendNumber(1);
    	       		psTemplateManager.AppendString(this.boltTypesComboBox.SelectedItem.ToString());
    	        	psTemplateManager.AppendString(this.boltDiametersComboBox.SelectedItem.ToString());
    	        	psTemplateManager.AppendString(this.boltAdditionalLengthTextBox.Text);
    	        	psTemplateManager.AppendString(this.anchorNameComboBox.Text.ToString());
    	        	psTemplateManager.AppendString(this.anchorDiametersComboBox.SelectedItem.ToString());
    	        	psTemplateManager.AppendString(this.anchorEmbedmentTextBox.Text);
    	        	psTemplateManager.AppendString(this.partThicknessTextBox.Text);
    	        	psTemplateManager.AppendString(this.anchorLengthComboBox.Text.ToString());
    	        	psTemplateManager.AppendBoolean(this.anchorLengthCheckBox.Checked);
    	        	psTemplateManager.SaveTemplateEntry("GSFBolt", 1);
    	        	BoltControl.CloseForm();
    	        	this.CreateSingleAnchor(anchorName, anchorDiameter, embedment, partThickness, anchorLength);
    	        	BoltControl.ShowForm(m_addIn);
    	    	}
    	    	else
    	    	{
    	    		string message = "Please enter the proper values"; 
    				string caption = "GSF TC Bolt/Exp. Anchor";
    				DialogResult result;
    				result = MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
    	    	}
    	    }
    	    else
    	    {
    	    	GoFindAnchor();
    	    }
        }

        private void CreateSingleAnchor(string AnchorName, double AnchorDiameter, double Embedment, double PartThickness, double AnchorLength)
        {
        	AnchorBolt anchorBolt = new AnchorBolt();
        	anchorBolt.SetParams(AnchorName, AnchorDiameter, Embedment, PartThickness, AnchorLength);

            PsPoint psP1 = new PsPoint();
            PsPoint psP2 = new PsPoint();
            PsPoint psP3 = new PsPoint();
            PsPoint reference = new PsPoint();
            
            PsPoint psPoint = new PsPoint();
            PsMatrix psMatrix = new PsMatrix();
            PsSelection psSelection = new PsSelection();
            PsSelection arg1 = psSelection;
            int num1 = arg1.PickPoint("Select insert point...", CoordSystem.kWcs, psP1);
            if (num1 == 1)
            {
                reference = psMatrix.TransformPoint(psP1);
                PsSelection arg2 = psSelection;
                num1 = arg2.PickPointWithReference("Select direction...", CoordSystem.kWcs, reference, psP2);
                if (num1 == 1)
                {
                	objId = anchorBolt.CreateAnchorCode111(psP1, psP2);
                }
            }
        }

        private void LengthCheckBox_CheckStateChanged(object sender, EventArgs e)
    	{
    		this.anchorLengthComboBox.Enabled = (this.anchorLengthCheckBox.CheckState == CheckState.Checked);
    		this.anchorNameComboBox.Enabled = (this.anchorLengthCheckBox.CheckState == CheckState.Checked);
    		PsTemplateManager psTemplateManager = new PsTemplateManager();
    		if (!this.anchorLengthCheckBox.Checked)
    		{
    			string oldDiameterValue = this.anchorDiametersComboBox.Text.ToString();
    			psTemplateManager.AppendNumber(1);
    			psTemplateManager.AppendString(this.boltTypesComboBox.SelectedItem.ToString());
    			psTemplateManager.AppendString(this.boltDiametersComboBox.SelectedItem.ToString());
    			psTemplateManager.AppendString(this.boltAdditionalLengthTextBox.Text);
    			psTemplateManager.AppendString(this.anchorNameComboBox.Text.ToString());
    			psTemplateManager.AppendString(this.anchorDiametersComboBox.SelectedItem.ToString());
    			psTemplateManager.AppendString(this.anchorEmbedmentTextBox.Text);
    			psTemplateManager.AppendString(this.partThicknessTextBox.Text);
    			psTemplateManager.AppendString(this.anchorLengthComboBox.Text.ToString());
    			psTemplateManager.AppendBoolean(this.anchorLengthCheckBox.Checked);
    			psTemplateManager.SaveTemplateEntry("GSFBolt", 1);
    			this.anchorDiametersComboBox.Items.Clear();
    			this.anchorLengthComboBox.Items.Clear();
    			this.anchorNameComboBox.Text = String.Empty;
    			this.anchorLengthComboBox.Text = String.Empty;
    			anchorboltdiameters_imperial = new string[]{
    	    		"0:0 1/4","0:0 3/8","0:0 1/2","0:0 5/8", "0:0 3/4", "0:0 7/8", "0:1", "0:1 1/8", "0:1 1/4"
    	    	};
    			this.anchorDiametersComboBox.Items.AddRange(anchorboltdiameters_imperial);
    			this.anchorDiametersComboBox.Text = oldDiameterValue;
    		}
    		else
    		{
    			psTemplateManager.LoadTemplateEntry("GSFBolt");
        		if (psTemplateManager.IsLoaded)
            	{
    		        long num1 = (long)psTemplateManager.get_Number(0);
    		        this.anchorNameComboBox.Text = psTemplateManager.get_String(3);
    		        this.anchorDiametersComboBox.Text = psTemplateManager.get_String(4);
    		        this.anchorLengthComboBox.Text = psTemplateManager.get_String(7);     
            	}
        		else
            	{
    		        this.anchorNameComboBox.Text = "EXP. ANCHOR";
    		        this.anchorDiametersComboBox.Text = anchorboltdiameters_imperial[0];
    		        this.anchorEmbedmentTextBox.Text = "0:0";
    		        this.partThicknessTextBox.Text = "0:0";
    		        this.anchorLengthComboBox.Text = "0:0";
    		        this.anchorLengthCheckBox.Checked = false;
           		}
    		}
    	}

        private void AnchorNameComboBox_SelectedIndexChanged(object sender, EventArgs e)
    	{
    		if (this.anchorLengthCheckBox.Checked)
    		{
    			PsUnits psUnits = new PsUnits();
    			string oldDiameterValue = this.anchorDiametersComboBox.Text.ToString();
    			string oldLengthValue = this.anchorLengthComboBox.Text.ToString();
    			this.anchorDiametersComboBox.Items.Clear();
    			this.anchorLengthComboBox.Items.Clear();
    			string[] arr = BoltDb.TablesToArray();
    			int pos = Array.IndexOf(arr, this.anchorNameComboBox.Text.ToString());
    			//index of "Type or select from the list" is 0.
    			if (pos < 1)
    			{
    				anchorboltdiameters_imperial = new string[]{
    	    		"0:0 1/4","0:0 3/8","0:0 1/2","0:0 5/8", "0:0 3/4", "0:0 7/8", "0:1", "0:1 1/8", "0:1 1/4"
    	    	};
    				this.anchorDiametersComboBox.Items.AddRange(anchorboltdiameters_imperial);
    				this.anchorDiametersComboBox.Text = oldDiameterValue;
    				this.anchorLengthComboBox.Text = oldLengthValue;
    			}
    			else
    			{
    				anchorboltdiameters_imperial = BoltDb.PopulateComboBoxWithDiameters(
    					this.anchorNameComboBox.Text.ToString()
    					);
    				this.anchorDiametersComboBox.Items.AddRange(anchorboltdiameters_imperial);
    				this.anchorDiametersComboBox.Text = anchorboltdiameters_imperial[0];
    				double criteria = psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text);
    				anchorboltlengths_imperial = BoltDb.PopulateComboBoxWithLengths(
    					this.anchorNameComboBox.SelectedItem.ToString(),
    					criteria
    					);
    				this.anchorLengthComboBox.Items.AddRange(anchorboltlengths_imperial);
    				this.anchorLengthComboBox.Text = anchorboltlengths_imperial[0];
    			}
    		}
    	}

        private void AnchorDiameterComboBox_SelectedIndexChanged(object sender, EventArgs e)
    	{
    		if (this.anchorLengthCheckBox.Checked)
    		{
    			PsUnits psUnits = new PsUnits();
    			string oldLengthValue = this.anchorLengthComboBox.Text.ToString();
    			this.anchorLengthComboBox.Items.Clear();
    			string[] arr = BoltDb.TablesToArray();
    			int pos = Array.IndexOf(arr, this.anchorNameComboBox.Text.ToString());
    			if (pos > 0)
    			{
    				double criteria = psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text);
    				anchorboltlengths_imperial = BoltDb.PopulateComboBoxWithLengths(
    					this.anchorNameComboBox.SelectedItem.ToString(),
    					criteria
    					);
    				this.anchorLengthComboBox.Items.AddRange(anchorboltlengths_imperial);
    				this.anchorLengthComboBox.Text = anchorboltlengths_imperial[0];
    			}
    			else
    			{
    				this.anchorLengthComboBox.Text = oldLengthValue;
    			}
    		}
    	}

        private void AnchorEmbedmentValueVerify(object sender, EventArgs e)
    	{
    		anchorEmbedmentValueVerified = anchorEmbedmentValuesVerified();
    	}

        private bool anchorEmbedmentValuesVerified()
    	{
    		if (this.anchorLengthCheckBox.Checked)
    		{
    			PsUnits psUnits = new PsUnits();
    			double length = psUnits.ConvertToNumeric(this.anchorLengthComboBox.Text);
    			double threadedLength;
    			double minEmbedment;
    			string[] arr = BoltDb.TablesToArray();
    			int pos = Array.IndexOf(arr, this.anchorNameComboBox.Text.ToString());
    			if (pos >= 1)
    			{
    				threadedLength = BoltDb.GetThreadedLength(
    					this.anchorNameComboBox.SelectedItem.ToString(),
    					psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text),
    					length
    					);
    				minEmbedment = BoltDb.GetMinEmbedment(
    					this.anchorNameComboBox.SelectedItem.ToString(),
    					psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text),
    					length
    					);
    			}
    			else
    			{
    				minEmbedment = 4*psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text);
    				threadedLength = length - 4*psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text);
    			}
    			double embedment = psUnits.ConvertToNumeric(this.anchorEmbedmentTextBox.Text);
    			AnchorBolt anchor = new AnchorBolt();
    			anchor.GetNutAndWasherSizesByDiameter(psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text));
    			double partThickness = psUnits.ConvertToNumeric(this.partThicknessTextBox.Text);
    			double e1 = length - partThickness - anchor.NutHeight - anchor.WasherThickness - 0.25D;
    			double e2 = length - threadedLength;
    			double emin = 0;
    			double emax = 0;
    			if (e1 > e2)
    			{
    				if (minEmbedment >= e2 && minEmbedment <= e1)
    				{
    					emax = e1;
    					emin = minEmbedment;
    					if (emax == emin)
    					{
    						if (embedment == emax)
    						{
    							this.anchorEmbedmentValuesRange.ForeColor = Color.Green;
    							this.anchorEmbedmentValuesRange.Text = "=> Is equal to" + Environment.NewLine + emax;
    							return true;
    						}
    						else
    						{
    							this.anchorEmbedmentValuesRange.ForeColor = Color.Red;
    							this.anchorEmbedmentValuesRange.Text = "=> Must be equal" + Environment.NewLine + "to " + emax;
    							return false;
    						}
    					}
    					else
    					{
    						if (embedment < emin)
    						{
    							this.anchorEmbedmentValuesRange.ForeColor = Color.Red;
    							this.anchorEmbedmentValuesRange.Text = "=> Out of range" + Environment.NewLine + "[" + emin + " - " + emax + "]";
    							return false;
    						}
    						else if (embedment > emax)
    						{
    							this.anchorEmbedmentValuesRange.ForeColor = Color.Red;
    							this.anchorEmbedmentValuesRange.Text = "=> Out of range" + Environment.NewLine + "[" + emin + " - " + emax + "]";
    							return false;
    						}
    						else
    						{
    							this.anchorEmbedmentValuesRange.ForeColor = Color.Green;
    							this.anchorEmbedmentValuesRange.Text = "=> In range" + Environment.NewLine + "[" + emin + " - " + emax + "]";
    							return true;
    						}
    					}
    				}
    				else if (minEmbedment > e1)
    				{
    					this.anchorEmbedmentValuesRange.ForeColor = Color.Red;
    					this.anchorEmbedmentValuesRange.Text = "Anchor is too" + Environment.NewLine + "short";
    					return false;
    				}
    				else
    				{
    					emax = e1;
    					emin = e2;
    					if (embedment < emin)
    					{
    						this.anchorEmbedmentValuesRange.ForeColor = Color.Red;
    						this.anchorEmbedmentValuesRange.Text = "=> Out of range" + Environment.NewLine + "[" + emin + " - " + emax + "]";
    						return false;
    					}
    					else if (embedment > emax)
    					{
    						this.anchorEmbedmentValuesRange.ForeColor = Color.Red;
    						this.anchorEmbedmentValuesRange.Text = "=> Out of range" + Environment.NewLine + "[" + emin + " - " + emax + "]";
    						return false;
    					}
    					else
    					{
    						this.anchorEmbedmentValuesRange.ForeColor = Color.Green;
    						this.anchorEmbedmentValuesRange.Text = "=> In range" + Environment.NewLine + "[" + emin + " - " + emax + "]";
    						return true;
    					}
    				}
    			}
    			else if (e1 < e2)
    			{
    				this.anchorEmbedmentValuesRange.ForeColor = Color.Red;
    				this.anchorEmbedmentValuesRange.Text = "Anchor is too" + Environment.NewLine + "short";
    				return false;
    			}
    			else
    			{
    				if (minEmbedment > e1)
    				{
    					this.anchorEmbedmentValuesRange.ForeColor = Color.Red;
    					this.anchorEmbedmentValuesRange.Text = "Anchor is too" + Environment.NewLine + "short";
    					return false;
    				}
    				else
    				{
    					emax = e1;
    					emin = e2;
    					if (embedment == emax)
    					{
    						this.anchorEmbedmentValuesRange.ForeColor = Color.Green;
    						this.anchorEmbedmentValuesRange.Text = "=> Is equal to" + Environment.NewLine + emax;
    						return true;
    					}
    					else
    					{
    						this.anchorEmbedmentValuesRange.ForeColor = Color.Red;
    						this.anchorEmbedmentValuesRange.Text = "=> Must be equal" + Environment.NewLine + "to " + emax;
    						return false;
    					}
    				}
    			}
    		}
    		else
    		{
    			this.anchorEmbedmentValuesRange.ForeColor = Color.Green;
    			this.anchorEmbedmentValuesRange.Text = "Searching" + Environment.NewLine +
    													"anchor" + Environment.NewLine +
    													"in database";
    			return true;
    		}
    	}

        private void cmdHelpBolt(object sender, System.EventArgs e)
        {
            Help();
        }

        private void Help()
    	{
    		string message = 
    		"ProSteel Addin GSF TC Bolt/Exp. Anchor v1.1" + Environment.NewLine +
    		"1. Creates tension control bolts" + Environment.NewLine +
    		"2. Replaces standard bolts with tension control bolts" + Environment.NewLine +
    		"3. Create expansion anchors with custom name" + Environment.NewLine +
    		"-------------------------------------------------------" + Environment.NewLine +
    		"Developed by Steel Detailer Albert E Sharapov" + Environment.NewLine +
    		"Copyright \u00A9 2017 Garrison Steel Fabricators, Inc." + Environment.NewLine +
    		"All Rights Reserved.";
    		string caption = "About GSF TC Bolt/Exp. Anchor";
    		DialogResult result;
    		result = MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
    	}

        private void cmdOk(object sender, System.EventArgs e)
        {
            Ok();
        }

        private void Ok()
    	{
    		BoltControl.CloseForm();
    	}

        private void cmdCancel(object sender, System.EventArgs e)
        {
            Cancel();
        }

        private void Cancel()
    	{
    		PsTransaction psTransaction = new PsTransaction();
    		psTransaction.EraseLongId((long)objId);
    		psTransaction.Close();
    		BoltControl.CloseForm();
    	}

        private void cmdAnchorData(object sender, System.EventArgs e)
        {
            AnchorData();
        }

        private void AnchorData()
    	{
    		PsUnits psUnits = new PsUnits();
    		string message;
    		string weblink = "Website: ";
    		string techdata = "Technical Data: ";
    		double Em;
    		double Tl;
    		
    		form1 = new System.Windows.Forms.Form();
    		form1.Text ="GSF TC Bolt/Exp. Anchor";

    		System.Windows.Forms.Panel panel1 = new System.Windows.Forms.Panel();
    		panel1.Location = new System.Drawing.Point(20, 20);
    		panel1.AutoSize = true;
    		panel1.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    		form1.Controls.Add(panel1);

    		System.Windows.Forms.Label label1 = new System.Windows.Forms.Label();

    		websitelinklabel = new System.Windows.Forms.LinkLabel();
    		techdatalinklabel = new System.Windows.Forms.LinkLabel();

    		string[] arr = BoltDb.TablesToArray();
    		int pos = Array.IndexOf(arr, this.anchorNameComboBox.Text.ToString());
    		if (pos >= 1)
    		{
    			Tl = BoltDb.GetThreadedLength(
    				this.anchorNameComboBox.SelectedItem.ToString(),
    				psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text),
    				psUnits.ConvertToNumeric(this.anchorLengthComboBox.Text)
    				);
    			Em = BoltDb.GetMinEmbedment(
    				this.anchorNameComboBox.SelectedItem.ToString(),
    				psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text),
    				psUnits.ConvertToNumeric(this.anchorLengthComboBox.Text)
    				);
    			message = 	"Minimum Embedment = " + psUnits.ConvertToText(Em) + Environment.NewLine +
    						Environment.NewLine +
    						"Threaded Length = " + psUnits.ConvertToText(Tl);
    			label1.Text = message;
    			label1.AutoSize = true;
    			panel1.Controls.Add(label1);

    			string[] weblinks = BoltDb.GetLinks(this.anchorNameComboBox.Text.ToString(), "LinkToWebsite");
    			string[] techdatalinks = BoltDb.GetLinks(this.anchorNameComboBox.Text.ToString(), "LinkToTechData");
    			int i = 1;
    			websitelinklabel.Text = weblink;
    			int j = websitelinklabel.Text.Length;
    			foreach(string s in weblinks)
    			{
    				websitelinklabel.Text = websitelinklabel.Text + " Link" + i.ToString() + " ";
    				websitelinklabel.Links.Add(j+1, 5, s);
    				j = websitelinklabel.Text.Length;
    				i++;
    			}
    			i = 1;
    			techdatalinklabel.Text = techdata;
    			j = techdatalinklabel.Text.Length;
    			foreach(string s in techdatalinks)
    			{
    				techdatalinklabel.Text = techdatalinklabel.Text + " Link" + i.ToString() + " ";
    				techdatalinklabel.Links.Add(j+1, 5, s);
    				j = techdatalinklabel.Text.Length;
    				i++;
    			}
    			websitelinklabel.Location = new System.Drawing.Point(
    				label1.Location.X,
    				label1.Bottom + 10);
    			websitelinklabel.AutoSize = true;
    			websitelinklabel.LinkClicked += new LinkLabelLinkClickedEventHandler(this.weblinkLabel_LinkClicked);
    			panel1.Controls.Add(websitelinklabel);
    			techdatalinklabel.Location = new System.Drawing.Point(
    				label1.Location.X,
    				websitelinklabel.Bottom + 10);
    			techdatalinklabel.AutoSize = true;
    			techdatalinklabel.LinkClicked += new LinkLabelLinkClickedEventHandler(this.techdatalinkLabel_LinkClicked);
    			panel1.Controls.Add(techdatalinklabel);
    		}
    		else
    		{
    			Em = 4*psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text);
    			Tl = psUnits.ConvertToNumeric(this.anchorLengthComboBox.Text) - 4*psUnits.ConvertToNumeric(this.anchorDiametersComboBox.Text);
    			message = 		"For anchors defined by user:" + Environment.NewLine +
    							Environment.NewLine +
    							"Minimum Embedment is assumed to be equal to 4 diameters," + Environment.NewLine +
    							"Threaded Length - anchor length minus minimum embedment." + Environment.NewLine +
    							Environment.NewLine +
    							"Minimum Embedment = " + psUnits.ConvertToText(Em) + Environment.NewLine +
    							Environment.NewLine +
    							"Threaded Length = " + psUnits.ConvertToText(Tl);
    			label1.Text = message;
    			label1.AutoSize = true;
    			panel1.Controls.Add(label1);
    		}
    		System.Windows.Forms.Button button1 = new System.Windows.Forms.Button();
    		button1.Location = new System.Drawing.Point(
    			panel1.Location.X + panel1.Size.Width/2 - button1.Size.Width/2,
    			panel1.Bottom + 15
    			);
    		button1.Text = "OK";
    		button1.Click += new EventHandler(this.CloseForm1);
    		form1.Controls.Add(button1);

    		form1.Size = new System.Drawing.Size(
    			panel1.Size.Width + 40,
    			20 + panel1.Size.Height + 15 + button1.Size.Height + 40
    			);

    		form1.FormBorderStyle = FormBorderStyle.FixedDialog;
    		form1.ShowIcon = false;
    		form1.MaximizeBox = false;
    		form1.MinimizeBox = false;
    		form1.StartPosition = FormStartPosition.CenterScreen;
    		form1.ShowDialog();
    	}

        private void weblinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    	{
    		websitelinklabel.Links[websitelinklabel.Links.IndexOf(e.Link)].Visited = true;
    		string target = e.Link.LinkData.ToString();
    		if (null != target && target.StartsWith("http://"))
    		{
    			System.Diagnostics.Process.Start(target);
    		}
    		else if (null != target && target.StartsWith("ftp://"))
    		{
    			System.Diagnostics.Process.Start(target);
    		}
    		else if (null != target && target.StartsWith("https://"))
    		{
    			System.Diagnostics.Process.Start(target);
    		}
    		else if (null != target && target.StartsWith("www"))
    		{
    			System.Diagnostics.Process.Start(target);
    		}
    		else
    		{
    			MessageBox.Show("incorrect URL or null", "GSF TC Bolt/Exp. Anchor", MessageBoxButtons.OK, MessageBoxIcon.Error);
    		}
    	}

        private void techdatalinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
    	{
    		techdatalinklabel.Links[techdatalinklabel.Links.IndexOf(e.Link)].Visited = true;
    		string target = e.Link.LinkData.ToString();
    		if (null != target && target.StartsWith("http://"))
    		{
    			System.Diagnostics.Process.Start(target);
    		}
    		else if (null != target && target.StartsWith("https://"))
    		{
    			System.Diagnostics.Process.Start(target);
    		}
    		else if (null != target && target.StartsWith("ftp://"))
    		{
    			System.Diagnostics.Process.Start(target);
    		}
    		else if (null != target && target.StartsWith("www"))
    		{
    			System.Diagnostics.Process.Start(target);
    		}
    		else
    		{
    			MessageBox.Show("incorrect URL or null", "GSF TC Bolt/Exp. Anchor", MessageBoxButtons.OK, MessageBoxIcon.Error);
    		}
    	}

        private void CloseForm1(object sender, EventArgs e)
    	{
    		form1.Close();
    	}

        private void GoFindAnchor()
    	{
    		this.anchorLengthCheckBox.Checked = true;
    	}

        private int ConvertBoolToInt(bool Checked)
    	{
    		int result = Checked ? 1: 0;
    		return result;
    	}
    }   
}   
