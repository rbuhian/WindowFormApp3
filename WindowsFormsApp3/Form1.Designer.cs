namespace WindowsFormsApp3
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.cmdPrint = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.cmdPrintToFile = new System.Windows.Forms.Button();
            this.cmdSelect = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.cmdConvert2DXF = new System.Windows.Forms.Button();
            this.cmdProjectProperties = new System.Windows.Forms.Button();
            this.cmdMiscList = new System.Windows.Forms.Button();
            this.cmdPDFSize = new System.Windows.Forms.Button();
            this.cmdMoveFile = new System.Windows.Forms.Button();
            this.cmdRemovePlates = new System.Windows.Forms.Button();
            this.cmdTestStack = new System.Windows.Forms.Button();
            this.cmdSendMessage = new System.Windows.Forms.Button();
            this.cmdShowWindowAsync = new System.Windows.Forms.Button();
            this.cmdFilter = new System.Windows.Forms.Button();
            this.cmdSegregateNC1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(57, 33);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(873, 20);
            this.textBox1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(372, 103);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(97, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Move NC1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(30, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "From";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 66);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(20, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "To";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(57, 63);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(873, 20);
            this.textBox2.TabIndex = 3;
            // 
            // cmdPrint
            // 
            this.cmdPrint.Location = new System.Drawing.Point(236, 103);
            this.cmdPrint.Name = "cmdPrint";
            this.cmdPrint.Size = new System.Drawing.Size(130, 23);
            this.cmdPrint.TabIndex = 5;
            this.cmdPrint.Text = "Print Drawing";
            this.cmdPrint.UseVisualStyleBackColor = true;
            this.cmdPrint.Click += new System.EventHandler(this.cmdPrint_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(83, 103);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(130, 23);
            this.button2.TabIndex = 6;
            this.button2.Text = "Get Env Variable";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(497, 103);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(130, 23);
            this.button3.TabIndex = 7;
            this.button3.Text = "Get Advance Option";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // cmdPrintToFile
            // 
            this.cmdPrintToFile.Location = new System.Drawing.Point(659, 103);
            this.cmdPrintToFile.Name = "cmdPrintToFile";
            this.cmdPrintToFile.Size = new System.Drawing.Size(130, 23);
            this.cmdPrintToFile.TabIndex = 8;
            this.cmdPrintToFile.Text = "Print To File";
            this.cmdPrintToFile.UseVisualStyleBackColor = true;
            this.cmdPrintToFile.Click += new System.EventHandler(this.cmdPrintToFile_Click);
            // 
            // cmdSelect
            // 
            this.cmdSelect.Location = new System.Drawing.Point(800, 103);
            this.cmdSelect.Name = "cmdSelect";
            this.cmdSelect.Size = new System.Drawing.Size(130, 23);
            this.cmdSelect.TabIndex = 9;
            this.cmdSelect.Text = "Select Drawing";
            this.cmdSelect.UseVisualStyleBackColor = true;
            this.cmdSelect.Click += new System.EventHandler(this.cmdSelect_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(83, 154);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(130, 23);
            this.button4.TabIndex = 10;
            this.button4.Text = "Create directory";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(236, 154);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(130, 23);
            this.button5.TabIndex = 11;
            this.button5.Text = "Directory Info";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // cmdConvert2DXF
            // 
            this.cmdConvert2DXF.Location = new System.Drawing.Point(372, 154);
            this.cmdConvert2DXF.Name = "cmdConvert2DXF";
            this.cmdConvert2DXF.Size = new System.Drawing.Size(130, 23);
            this.cmdConvert2DXF.TabIndex = 12;
            this.cmdConvert2DXF.Text = "Convert to DXF";
            this.cmdConvert2DXF.UseVisualStyleBackColor = true;
            this.cmdConvert2DXF.Click += new System.EventHandler(this.cmdConvert2DXF_Click);
            // 
            // cmdProjectProperties
            // 
            this.cmdProjectProperties.Location = new System.Drawing.Point(508, 154);
            this.cmdProjectProperties.Name = "cmdProjectProperties";
            this.cmdProjectProperties.Size = new System.Drawing.Size(130, 23);
            this.cmdProjectProperties.TabIndex = 13;
            this.cmdProjectProperties.Text = "Project Properties";
            this.cmdProjectProperties.UseVisualStyleBackColor = true;
            this.cmdProjectProperties.Click += new System.EventHandler(this.cmdProjectProperties_Click);
            // 
            // cmdMiscList
            // 
            this.cmdMiscList.Location = new System.Drawing.Point(659, 154);
            this.cmdMiscList.Name = "cmdMiscList";
            this.cmdMiscList.Size = new System.Drawing.Size(130, 23);
            this.cmdMiscList.TabIndex = 14;
            this.cmdMiscList.Text = "Misc List";
            this.cmdMiscList.UseVisualStyleBackColor = true;
            this.cmdMiscList.Click += new System.EventHandler(this.cmdMiscList_Click);
            // 
            // cmdPDFSize
            // 
            this.cmdPDFSize.Location = new System.Drawing.Point(800, 154);
            this.cmdPDFSize.Name = "cmdPDFSize";
            this.cmdPDFSize.Size = new System.Drawing.Size(130, 23);
            this.cmdPDFSize.TabIndex = 15;
            this.cmdPDFSize.Text = "PDF Size";
            this.cmdPDFSize.UseVisualStyleBackColor = true;
            this.cmdPDFSize.Click += new System.EventHandler(this.cmdPDFSize_Click);
            // 
            // cmdMoveFile
            // 
            this.cmdMoveFile.Location = new System.Drawing.Point(83, 210);
            this.cmdMoveFile.Name = "cmdMoveFile";
            this.cmdMoveFile.Size = new System.Drawing.Size(130, 23);
            this.cmdMoveFile.TabIndex = 16;
            this.cmdMoveFile.Text = "Move File";
            this.cmdMoveFile.UseVisualStyleBackColor = true;
            this.cmdMoveFile.Click += new System.EventHandler(this.cmdMoveFile_Click);
            // 
            // cmdRemovePlates
            // 
            this.cmdRemovePlates.Location = new System.Drawing.Point(236, 210);
            this.cmdRemovePlates.Name = "cmdRemovePlates";
            this.cmdRemovePlates.Size = new System.Drawing.Size(130, 23);
            this.cmdRemovePlates.TabIndex = 17;
            this.cmdRemovePlates.Text = "Remove Plates ";
            this.cmdRemovePlates.UseVisualStyleBackColor = true;
            this.cmdRemovePlates.Click += new System.EventHandler(this.cmdRemovePlates_Click);
            // 
            // cmdTestStack
            // 
            this.cmdTestStack.Location = new System.Drawing.Point(372, 210);
            this.cmdTestStack.Name = "cmdTestStack";
            this.cmdTestStack.Size = new System.Drawing.Size(130, 23);
            this.cmdTestStack.TabIndex = 18;
            this.cmdTestStack.Text = "Test Stack";
            this.cmdTestStack.UseVisualStyleBackColor = true;
            this.cmdTestStack.Click += new System.EventHandler(this.cmdTestStack_Click);
            // 
            // cmdSendMessage
            // 
            this.cmdSendMessage.Location = new System.Drawing.Point(508, 210);
            this.cmdSendMessage.Name = "cmdSendMessage";
            this.cmdSendMessage.Size = new System.Drawing.Size(130, 23);
            this.cmdSendMessage.TabIndex = 19;
            this.cmdSendMessage.Text = "Send Message";
            this.cmdSendMessage.UseVisualStyleBackColor = true;
            this.cmdSendMessage.Click += new System.EventHandler(this.cmdSendMessage_Click);
            // 
            // cmdShowWindowAsync
            // 
            this.cmdShowWindowAsync.Location = new System.Drawing.Point(659, 210);
            this.cmdShowWindowAsync.Name = "cmdShowWindowAsync";
            this.cmdShowWindowAsync.Size = new System.Drawing.Size(130, 23);
            this.cmdShowWindowAsync.TabIndex = 20;
            this.cmdShowWindowAsync.Text = "ShowWindowAsync";
            this.cmdShowWindowAsync.UseVisualStyleBackColor = true;
            this.cmdShowWindowAsync.Click += new System.EventHandler(this.cmdShowWindowAsync_Click);
            // 
            // cmdFilter
            // 
            this.cmdFilter.Location = new System.Drawing.Point(800, 210);
            this.cmdFilter.Name = "cmdFilter";
            this.cmdFilter.Size = new System.Drawing.Size(130, 23);
            this.cmdFilter.TabIndex = 21;
            this.cmdFilter.Text = "Filter by UserIssue";
            this.cmdFilter.UseVisualStyleBackColor = true;
            this.cmdFilter.Click += new System.EventHandler(this.cmdFilter_Click);
            // 
            // cmdSegregateNC1
            // 
            this.cmdSegregateNC1.Location = new System.Drawing.Point(83, 257);
            this.cmdSegregateNC1.Name = "cmdSegregateNC1";
            this.cmdSegregateNC1.Size = new System.Drawing.Size(130, 23);
            this.cmdSegregateNC1.TabIndex = 22;
            this.cmdSegregateNC1.Text = "Segregate NC1";
            this.cmdSegregateNC1.UseVisualStyleBackColor = true;
            this.cmdSegregateNC1.Click += new System.EventHandler(this.cmdSegregateNC1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(968, 328);
            this.Controls.Add(this.cmdSegregateNC1);
            this.Controls.Add(this.cmdFilter);
            this.Controls.Add(this.cmdShowWindowAsync);
            this.Controls.Add(this.cmdSendMessage);
            this.Controls.Add(this.cmdTestStack);
            this.Controls.Add(this.cmdRemovePlates);
            this.Controls.Add(this.cmdMoveFile);
            this.Controls.Add(this.cmdPDFSize);
            this.Controls.Add(this.cmdMiscList);
            this.Controls.Add(this.cmdProjectProperties);
            this.Controls.Add(this.cmdConvert2DXF);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.cmdSelect);
            this.Controls.Add(this.cmdPrintToFile);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.cmdPrint);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button cmdPrint;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button cmdPrintToFile;
        private System.Windows.Forms.Button cmdSelect;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button cmdConvert2DXF;
        private System.Windows.Forms.Button cmdProjectProperties;
        private System.Windows.Forms.Button cmdMiscList;
        private System.Windows.Forms.Button cmdPDFSize;
        private System.Windows.Forms.Button cmdMoveFile;
        private System.Windows.Forms.Button cmdRemovePlates;
        private System.Windows.Forms.Button cmdTestStack;
        private System.Windows.Forms.Button cmdSendMessage;
        private System.Windows.Forms.Button cmdShowWindowAsync;
        private System.Windows.Forms.Button cmdFilter;
        private System.Windows.Forms.Button cmdSegregateNC1;
    }
}

