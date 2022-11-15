namespace CS_Excel
{
   partial class Form1
   {
      /// <summary>
      /// Обязательная переменная конструктора.
      /// </summary>
      private System.ComponentModel.IContainer components = null;

      /// <summary>
      /// Освободить все используемые ресурсы.
      /// </summary>
      /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
      protected override void Dispose(bool disposing)
      {
         if (disposing && (components != null))
         {
            components.Dispose();
         }
         base.Dispose(disposing);
      }

      #region Код, автоматически созданный конструктором форм Windows

      /// <summary>
      /// Требуемый метод для поддержки конструктора — не изменяйте 
      /// содержимое этого метода с помощью редактора кода.
      /// </summary>
      private void InitializeComponent()
      {
         this.dataGridView1 = new System.Windows.Forms.DataGridView();
         this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
         this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
         this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
         this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
         this.button1 = new System.Windows.Forms.Button();
         this.button2 = new System.Windows.Forms.Button();
         this.button3 = new System.Windows.Forms.Button();
         this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
         this.textBox1 = new System.Windows.Forms.TextBox();
         this.textBox2 = new System.Windows.Forms.TextBox();
         this.label1 = new System.Windows.Forms.Label();
         this.label2 = new System.Windows.Forms.Label();
         this.textBox3 = new System.Windows.Forms.TextBox();
         this.textBox4 = new System.Windows.Forms.TextBox();
         this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
         this.printDialog1 = new System.Windows.Forms.PrintDialog();
         this.panel1 = new System.Windows.Forms.Panel();
         ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
         this.panel1.SuspendLayout();
         this.SuspendLayout();
         // 
         // dataGridView1
         // 
         this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
         this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4});
         this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
         this.dataGridView1.Location = new System.Drawing.Point(0, 53);
         this.dataGridView1.Name = "dataGridView1";
         this.dataGridView1.Size = new System.Drawing.Size(844, 397);
         this.dataGridView1.TabIndex = 0;
         // 
         // Column1
         // 
         this.Column1.HeaderText = "Column1";
         this.Column1.Name = "Column1";
         // 
         // Column2
         // 
         this.Column2.HeaderText = "Column2";
         this.Column2.Name = "Column2";
         // 
         // Column3
         // 
         this.Column3.HeaderText = "Column3";
         this.Column3.Name = "Column3";
         // 
         // Column4
         // 
         this.Column4.HeaderText = "Column4";
         this.Column4.Name = "Column4";
         // 
         // button1
         // 
         this.button1.Location = new System.Drawing.Point(15, 10);
         this.button1.Name = "button1";
         this.button1.Size = new System.Drawing.Size(75, 33);
         this.button1.TabIndex = 1;
         this.button1.Text = "add row";
         this.button1.UseVisualStyleBackColor = true;
         // 
         // button2
         // 
         this.button2.Location = new System.Drawing.Point(96, 10);
         this.button2.Name = "button2";
         this.button2.Size = new System.Drawing.Size(75, 33);
         this.button2.TabIndex = 2;
         this.button2.Text = "To Excel";
         this.button2.UseVisualStyleBackColor = true;
         this.button2.Click += new System.EventHandler(this.button2_Click);
         // 
         // button3
         // 
         this.button3.Location = new System.Drawing.Point(177, 10);
         this.button3.Name = "button3";
         this.button3.Size = new System.Drawing.Size(75, 33);
         this.button3.TabIndex = 3;
         this.button3.Text = "From Excel";
         this.button3.UseVisualStyleBackColor = true;
         this.button3.Click += new System.EventHandler(this.button3_Click);
         // 
         // openFileDialog1
         // 
         this.openFileDialog1.FileName = "openFileDialog1";
         // 
         // textBox1
         // 
         this.textBox1.Location = new System.Drawing.Point(628, 3);
         this.textBox1.Name = "textBox1";
         this.textBox1.Size = new System.Drawing.Size(72, 20);
         this.textBox1.TabIndex = 4;
         this.textBox1.Text = "2";
         // 
         // textBox2
         // 
         this.textBox2.Location = new System.Drawing.Point(628, 26);
         this.textBox2.Name = "textBox2";
         this.textBox2.Size = new System.Drawing.Size(72, 20);
         this.textBox2.TabIndex = 4;
         this.textBox2.Text = "1";
         // 
         // label1
         // 
         this.label1.AutoSize = true;
         this.label1.Location = new System.Drawing.Point(586, 7);
         this.label1.Name = "label1";
         this.label1.Size = new System.Drawing.Size(29, 13);
         this.label1.TabIndex = 5;
         this.label1.Text = "Row";
         // 
         // label2
         // 
         this.label2.AutoSize = true;
         this.label2.Location = new System.Drawing.Point(586, 30);
         this.label2.Name = "label2";
         this.label2.Size = new System.Drawing.Size(22, 13);
         this.label2.TabIndex = 6;
         this.label2.Text = "Col";
         // 
         // textBox3
         // 
         this.textBox3.Location = new System.Drawing.Point(704, 3);
         this.textBox3.Name = "textBox3";
         this.textBox3.Size = new System.Drawing.Size(72, 20);
         this.textBox3.TabIndex = 7;
         this.textBox3.Text = "3";
         // 
         // textBox4
         // 
         this.textBox4.Location = new System.Drawing.Point(704, 26);
         this.textBox4.Name = "textBox4";
         this.textBox4.Size = new System.Drawing.Size(72, 20);
         this.textBox4.TabIndex = 8;
         this.textBox4.Text = "4";
         // 
         // printDialog1
         // 
         this.printDialog1.UseEXDialog = true;
         // 
         // panel1
         // 
         this.panel1.Controls.Add(this.button3);
         this.panel1.Controls.Add(this.textBox4);
         this.panel1.Controls.Add(this.button1);
         this.panel1.Controls.Add(this.textBox3);
         this.panel1.Controls.Add(this.label2);
         this.panel1.Controls.Add(this.button2);
         this.panel1.Controls.Add(this.label1);
         this.panel1.Controls.Add(this.textBox1);
         this.panel1.Controls.Add(this.textBox2);
         this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
         this.panel1.Location = new System.Drawing.Point(0, 0);
         this.panel1.Name = "panel1";
         this.panel1.Size = new System.Drawing.Size(844, 53);
         this.panel1.TabIndex = 9;
         // 
         // Form1
         // 
         this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
         this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
         this.ClientSize = new System.Drawing.Size(844, 450);
         this.Controls.Add(this.dataGridView1);
         this.Controls.Add(this.panel1);
         this.Name = "Form1";
         this.Text = "Form1";
         ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
         this.panel1.ResumeLayout(false);
         this.panel1.PerformLayout();
         this.ResumeLayout(false);

      }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.PrintDialog printDialog1;
        private System.Windows.Forms.Panel panel1;
    }
}

