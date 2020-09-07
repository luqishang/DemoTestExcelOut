namespace TestExcelOutput
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnTestExcelOutput = new System.Windows.Forms.Button();
            this.saveExcelFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // btnTestExcelOutput
            // 
            this.btnTestExcelOutput.Location = new System.Drawing.Point(291, 310);
            this.btnTestExcelOutput.Name = "btnTestExcelOutput";
            this.btnTestExcelOutput.Size = new System.Drawing.Size(175, 53);
            this.btnTestExcelOutput.TabIndex = 0;
            this.btnTestExcelOutput.Text = "テスト帳票";
            this.btnTestExcelOutput.UseVisualStyleBackColor = true;
            this.btnTestExcelOutput.Click += new System.EventHandler(this.btnTestExcelOutput_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnTestExcelOutput);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnTestExcelOutput;
        private System.Windows.Forms.SaveFileDialog saveExcelFileDialog;
    }
}

