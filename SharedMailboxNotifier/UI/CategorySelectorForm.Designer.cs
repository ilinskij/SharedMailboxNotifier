namespace SharedMailboxNotifier.UI
{
    partial class CategorySelectorForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this._labelAvailable = new System.Windows.Forms.Label();
            this._labelSelected = new System.Windows.Forms.Label();
            this._listAvailable = new System.Windows.Forms.ListBox();
            this._listSelected = new System.Windows.Forms.ListBox();
            this._btnAdd = new System.Windows.Forms.Button();
            this._btnRemove = new System.Windows.Forms.Button();
            this._btnUp = new System.Windows.Forms.Button();
            this._btnDown = new System.Windows.Forms.Button();
            this._labelLimit = new System.Windows.Forms.Label();
            this._btnOk = new System.Windows.Forms.Button();
            this._btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // _labelAvailable
            // 
            this._labelAvailable.AutoSize = true;
            this._labelAvailable.Location = new System.Drawing.Point(12, 8);
            this._labelAvailable.Name = "_labelAvailable";
            this._labelAvailable.Size = new System.Drawing.Size(50, 13);
            this._labelAvailable.TabIndex = 0;
            this._labelAvailable.Text = "Available";
            // 
            // _labelSelected
            // 
            this._labelSelected.AutoSize = true;
            this._labelSelected.Location = new System.Drawing.Point(221, 8);
            this._labelSelected.Name = "_labelSelected";
            this._labelSelected.Size = new System.Drawing.Size(49, 13);
            this._labelSelected.TabIndex = 1;
            this._labelSelected.Text = "Selected";
            // 
            // _listAvailable
            // 
            this._listAvailable.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this._listAvailable.FormattingEnabled = true;
            this._listAvailable.IntegralHeight = false;
            this._listAvailable.Location = new System.Drawing.Point(12, 24);
            this._listAvailable.Name = "_listAvailable";
            this._listAvailable.Size = new System.Drawing.Size(170, 220);
            this._listAvailable.TabIndex = 2;
            this._listAvailable.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.OnDrawItem);
            this._listAvailable.SelectedIndexChanged += new System.EventHandler(this.OnSelectionChanged);
            this._listAvailable.DoubleClick += new System.EventHandler(this.OnAddClick);
            // 
            // _listSelected
            // 
            this._listSelected.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this._listSelected.FormattingEnabled = true;
            this._listSelected.IntegralHeight = false;
            this._listSelected.Location = new System.Drawing.Point(224, 24);
            this._listSelected.Name = "_listSelected";
            this._listSelected.Size = new System.Drawing.Size(170, 220);
            this._listSelected.TabIndex = 3;
            this._listSelected.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.OnDrawItem);
            this._listSelected.SelectedIndexChanged += new System.EventHandler(this.OnSelectionChanged);
            this._listSelected.DoubleClick += new System.EventHandler(this.OnRemoveClick);
            // 
            // _btnAdd
            // 
            this._btnAdd.Location = new System.Drawing.Point(188, 70);
            this._btnAdd.Name = "_btnAdd";
            this._btnAdd.Size = new System.Drawing.Size(30, 28);
            this._btnAdd.TabIndex = 4;
            this._btnAdd.Text = "→";
            this._btnAdd.UseVisualStyleBackColor = true;
            this._btnAdd.Click += new System.EventHandler(this.OnAddClick);
            // 
            // _btnRemove
            // 
            this._btnRemove.Location = new System.Drawing.Point(188, 104);
            this._btnRemove.Name = "_btnRemove";
            this._btnRemove.Size = new System.Drawing.Size(30, 28);
            this._btnRemove.TabIndex = 5;
            this._btnRemove.Text = "←";
            this._btnRemove.UseVisualStyleBackColor = true;
            this._btnRemove.Click += new System.EventHandler(this.OnRemoveClick);
            // 
            // _btnUp
            // 
            this._btnUp.Location = new System.Drawing.Point(188, 152);
            this._btnUp.Name = "_btnUp";
            this._btnUp.Size = new System.Drawing.Size(30, 28);
            this._btnUp.TabIndex = 6;
            this._btnUp.Text = "▲";
            this._btnUp.UseVisualStyleBackColor = true;
            this._btnUp.Click += new System.EventHandler(this.OnMoveUpClick);
            // 
            // _btnDown
            // 
            this._btnDown.Location = new System.Drawing.Point(188, 186);
            this._btnDown.Name = "_btnDown";
            this._btnDown.Size = new System.Drawing.Size(30, 28);
            this._btnDown.TabIndex = 7;
            this._btnDown.Text = "▼";
            this._btnDown.UseVisualStyleBackColor = true;
            this._btnDown.Click += new System.EventHandler(this.OnMoveDownClick);
            // 
            // _labelLimit
            // 
            this._labelLimit.AutoSize = true;
            this._labelLimit.ForeColor = System.Drawing.SystemColors.GrayText;
            this._labelLimit.Location = new System.Drawing.Point(10, 258);
            this._labelLimit.Name = "_labelLimit";
            this._labelLimit.Size = new System.Drawing.Size(82, 13);
            this._labelLimit.TabIndex = 8;
            this._labelLimit.Text = "Selected: 0 of 5";
            // 
            // _btnOk
            // 
            this._btnOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            this._btnOk.Location = new System.Drawing.Point(238, 255);
            this._btnOk.Name = "_btnOk";
            this._btnOk.Size = new System.Drawing.Size(75, 25);
            this._btnOk.TabIndex = 9;
            this._btnOk.Text = "OK";
            this._btnOk.UseVisualStyleBackColor = true;
            this._btnOk.Click += new System.EventHandler(this.OnOkClick);
            // 
            // _btnCancel
            // 
            this._btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this._btnCancel.Location = new System.Drawing.Point(319, 255);
            this._btnCancel.Name = "_btnCancel";
            this._btnCancel.Size = new System.Drawing.Size(75, 25);
            this._btnCancel.TabIndex = 10;
            this._btnCancel.Text = "Cancel";
            this._btnCancel.UseVisualStyleBackColor = true;
            // 
            // CategorySelectorForm
            // 
            this.AcceptButton = this._btnOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this._btnCancel;
            this.ClientSize = new System.Drawing.Size(406, 291);
            this.Controls.Add(this._labelAvailable);
            this.Controls.Add(this._labelSelected);
            this.Controls.Add(this._listAvailable);
            this.Controls.Add(this._listSelected);
            this.Controls.Add(this._btnAdd);
            this.Controls.Add(this._btnRemove);
            this.Controls.Add(this._btnUp);
            this.Controls.Add(this._btnDown);
            this.Controls.Add(this._labelLimit);
            this.Controls.Add(this._btnOk);
            this.Controls.Add(this._btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CategorySelectorForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Category Selection";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label _labelAvailable;
        private System.Windows.Forms.Label _labelSelected;
        private System.Windows.Forms.ListBox _listAvailable;
        private System.Windows.Forms.ListBox _listSelected;
        private System.Windows.Forms.Button _btnAdd;
        private System.Windows.Forms.Button _btnRemove;
        private System.Windows.Forms.Button _btnUp;
        private System.Windows.Forms.Button _btnDown;
        private System.Windows.Forms.Label _labelLimit;
        private System.Windows.Forms.Button _btnOk;
        private System.Windows.Forms.Button _btnCancel;
    }
}
