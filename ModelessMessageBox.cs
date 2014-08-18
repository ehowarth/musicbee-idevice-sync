using System;
using System.Drawing;
using System.Windows.Forms;

namespace MusicBeeDeviceSyncPlugin
{
	public sealed class ModelessMessageBox : Form
	{
		public static Form Show(Form owner, string text, string caption, string okText, string cancelText)
		{
			Form dialog = null;
			owner.Invoke((Action)(() => dialog = new ModelessMessageBox(owner, text, caption, okText, cancelText)));
			return dialog;
		}

		private ModelessMessageBox(Form owner, string text, string caption, string okText, string cancelText)
		{
			var okButton = new Button();
			okButton.DialogResult = DialogResult.OK;
			okButton.UseVisualStyleBackColor = true;
			okButton.AutoSize = true;
			okButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
			okButton.Text = okText;
			okButton.Click += ButtonClicked;

			var cancelButton = new Button();
			cancelButton.DialogResult = DialogResult.Cancel;
			cancelButton.UseVisualStyleBackColor = true;
			cancelButton.AutoSize = true;
			cancelButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
			cancelButton.Text = cancelText;
			cancelButton.Click += ButtonClicked;

			var buttonPanel = new FlowLayoutPanel();
			buttonPanel.SuspendLayout();
			buttonPanel.FlowDirection = FlowDirection.RightToLeft;
			buttonPanel.AutoSize = true;
			buttonPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
			buttonPanel.Dock = DockStyle.Bottom;
			buttonPanel.Controls.Add(cancelButton);
			buttonPanel.Controls.Add(okButton);

			var bodyText = new Label();
			bodyText.SuspendLayout();
			bodyText.AutoSize = true;
			bodyText.MaximumSize = new Size(400, 400);
			bodyText.Padding = new Padding(10);
			bodyText.Dock = DockStyle.Fill;
			bodyText.Text = text;

			this.SuspendLayout();
			this.AutoScaleMode = AutoScaleMode.Font;
			this.AutoSize = true;
			this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
			this.AcceptButton = okButton;
			this.CancelButton = cancelButton;
			this.ControlBox = false;
			this.Controls.Add(bodyText);
			this.Controls.Add(buttonPanel);
			this.FormBorderStyle = FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.ShowIcon = false;
			this.ShowInTaskbar = false;
			this.StartPosition = FormStartPosition.CenterScreen;
			this.Text = caption;

			buttonPanel.ResumeLayout(false);
			buttonPanel.PerformLayout();
			bodyText.ResumeLayout(false);
			bodyText.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

			owner.AddOwnedForm(this);
			this.Show(owner);
			this.Invalidate();
		}

		private void ButtonClicked(object button, EventArgs e)
		{
			this.DialogResult = ((Button)button).DialogResult;
			this.Close();
		}
	}
}
