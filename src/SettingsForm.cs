using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace AutoFrom
{
    public sealed class SettingsForm : Form
    {
        readonly CheckBox enabled = new CheckBox { Text = "Automatically select the sender", AutoSize = true };
        readonly DataGridView senders = Grid();
        readonly DataGridView rules = Grid();
        readonly ComboBox testAccount = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        readonly TextBox testRecipients = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical };
        readonly Label testResult = new Label { Dock = DockStyle.Fill, Text = "Enter all To, Cc and Bcc addresses to preview the result.", Padding = new Padding(0, 12, 0, 0) };
        readonly Action<Config> save;
        readonly string[] accounts;
        readonly TabControl tabs = new TabControl { Dock = DockStyle.Fill };
        public SettingsForm(Config initial, string[] profileAccounts, string startupError, Action<Config> onSave)
        {
            save = onSave;
            if (string.IsNullOrWhiteSpace(startupError)) startupError = null;
            initial = SettingsStore.Validate(initial ?? SettingsStore.Empty());
            accounts = profileAccounts.Concat(initial.Senders.Select(s => s.Account)).Distinct().OrderBy(a => a).ToArray();
            Text = "AutoFrom settings - 1.2.6 (alias pilot)";
            Font = new Font("Segoe UI", 10);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(1100, 650);
            MinimumSize = new Size(850, 580);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), RowCount = 5, ColumnCount = 1 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, startupError == null ? 44 : 70));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            layout.Controls.Add(new Label { Text = "Choose the right sender, automatically", AutoSize = true, Font = new Font(Font.FontFamily, 17, FontStyle.Bold) }, 0, 0);
            enabled.Checked = initial.Enabled;
            layout.Controls.Add(enabled, 0, 1);
            layout.Controls.Add(new Label { Dock = DockStyle.Fill, Text = startupError == null
                ? "To always use an address, add it under Senders and tick Use as default. Recipient rules are optional."
                : "Your saved settings could not be loaded. Configure them here and save to recover. The previous file will be backed up.\r\n" + startupError,
                ForeColor = startupError == null ? Color.DimGray : Color.DarkRed }, 0, 2);
            layout.Controls.Add(tabs, 0, 3);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, 12, 0, 0) };
            var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
            var apply = new Button { Text = "Save settings", AutoSize = true };
            apply.Click += SaveClicked;
            buttons.Controls.Add(cancel); buttons.Controls.Add(apply);
            layout.Controls.Add(buttons, 0, 4);
            CancelButton = cancel;
            Controls.Add(layout);

            senders.Columns.Add(TextColumn("id", "Label (any name)", 18));
            var accountColumn = new DataGridViewComboBoxColumn { Name = "account", HeaderText = "Owning Outlook account", FillWeight = 24, FlatStyle = FlatStyle.Flat };
            accountColumn.Items.AddRange(accounts);
            senders.Columns.Add(accountColumn);
            senders.Columns.Add(TextColumn("alias", "Alias address", 24));
            senders.Columns.Add(TextColumn("represented", "Shared / delegated mailbox", 24));
            senders.Columns.Add(new DataGridViewCheckBoxColumn { Name = "default", HeaderText = "Use as default", FillWeight = 10 });
            foreach (Sender s in initial.Senders) senders.Rows.Add(s.Id, s.Account, s.Alias, s.Represented, s.DefaultForAccount);
            AddGridTab("Senders", "Alias: select its owning account and enter the alias address; leave Shared / delegated mailbox blank. Tick Use as default to apply it to that account's drafts, even before recipients are entered. Alias sending requires Exchange Online support and a delivered-message test.", senders,
                () => senders.Rows.Add(NextName(senders, "sender"), accounts.FirstOrDefault(), "", "", false));

            rules.Columns.Add(TextColumn("id", "Rule name", 14));
            rules.Columns.Add(TextColumn("priority", "Priority", 8));
            rules.Columns.Add(new DataGridViewComboBoxColumn { Name = "sender", HeaderText = "Use sender", FillWeight = 14, FlatStyle = FlatStyle.Flat });
            rules.Columns.Add(TextColumn("recipients", "Exact recipients", 18));
            rules.Columns.Add(TextColumn("domains", "Recipient domains", 16));
            rules.Columns.Add(TextColumn("accounts", "Original accounts (optional)", 20));
            rules.Columns.Add(new DataGridViewCheckBoxColumn { Name = "subdomains", HeaderText = "Include subdomains", FillWeight = 10 });
            RefreshSenderChoices();
            foreach (Rule r in initial.Rules) rules.Rows.Add(r.Id, r.Priority, r.SenderId, string.Join("; ", r.Recipients), string.Join("; ", r.Domains), string.Join("; ", r.Accounts), r.IncludeSubdomains);
            AddGridTab("Rules", "Optional overrides for selected recipients. For one sender on all mail, use the Senders tab's Use as default checkbox instead. Separate entries with semicolons; lower numbers win per recipient. All recipients must agree on the same sender.", rules,
                () => { RefreshSenderChoices(); rules.Rows.Add(NextName(rules, "rule"), 100, SenderNames().FirstOrDefault(), "", "", "", false); });
            tabs.Selecting += (s, e) => { senders.EndEdit(); rules.EndEdit(); RefreshSenderChoices(); };
            AddTestTab();
        }
        static DataGridView Grid()
        {
            var grid = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersVisible = false,
                BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false, AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize };
            grid.DataError += (s, e) => {
                e.ThrowException = false;
                if (e.RowIndex >= 0) grid.Rows[e.RowIndex].ErrorText = "Choose an available value, then save again.";
            };
            return grid;
        }
        static DataGridViewTextBoxColumn TextColumn(string name, string title, float weight)
        {
            return new DataGridViewTextBoxColumn { Name = name, HeaderText = title, FillWeight = weight, MinimumWidth = 70,
                DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True } };
        }
        static string Value(DataGridViewRow row, string column) { return Convert.ToString(row.Cells[column].Value).Trim(); }
        static string[] List(string value) { return value.Split(new[] {';', ',', '\r', '\n'}, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).Where(s => s.Length > 0).ToArray(); }
        static string NextName(DataGridView grid, string prefix)
        {
            int i = 1;
            while (grid.Rows.Cast<DataGridViewRow>().Any(r => Value(r, "id") == prefix + "-" + i)) i++;
            return prefix + "-" + i;
        }
        IEnumerable<string> SenderNames() { return senders.Rows.Cast<DataGridViewRow>().Select(r => Value(r, "id")).Where(s => s.Length > 0).Distinct(); }
        void RefreshSenderChoices()
        {
            if (!rules.Columns.Contains("sender")) return;
            var column = (DataGridViewComboBoxColumn)rules.Columns["sender"];
            // Retain old references visibly; validation explains a renamed/deleted sender at Save.
            var names = SenderNames().Concat(rules.Rows.Cast<DataGridViewRow>().Select(r => Value(r, "sender"))).Where(s => s.Length > 0).Distinct().ToArray();
            // Clearing a populated combo column can trigger a synchronous DataError while
            // the grid paints existing cells. Append choices; dangling references fail Save.
            foreach (string name in names) if (!column.Items.Contains(name)) column.Items.Add(name);
            var valid = new HashSet<string>(SenderNames());
            foreach (DataGridViewRow row in rules.Rows)
                row.Cells["sender"].ErrorText = valid.Contains(Value(row, "sender")) ? "" : "This sender label no longer exists. Choose a current sender or remove this rule.";
        }
        void AddGridTab(string title, string help, DataGridView grid, Action add)
        {
            var tab = new TabPage(title) { Padding = new Padding(12) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            layout.Controls.Add(new Label { Text = help, Dock = DockStyle.Fill }, 0, 0);
            layout.Controls.Add(grid, 0, 1);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 6, 0, 0) };
            var addButton = new Button { Text = "Add " + (title == "Senders" ? "sender" : "rule"), AutoSize = true };
            var remove = new Button { Text = "Remove selected", AutoSize = true };
            addButton.Click += (s, e) => { grid.EndEdit(); add(); if (grid.Rows.Count > 0) grid.CurrentCell = grid.Rows[grid.Rows.Count - 1].Cells[0]; };
            remove.Click += (s, e) => { if (grid.CurrentRow != null) { grid.EndEdit(); grid.Rows.Remove(grid.CurrentRow); } };
            buttons.Controls.Add(addButton); buttons.Controls.Add(remove);
            layout.Controls.Add(buttons, 0, 2); tab.Controls.Add(layout); tabs.TabPages.Add(tab);
        }
        void AddTestTab()
        {
            var tab = new TabPage("Test rules") { Padding = new Padding(16) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 7, ColumnCount = 1 };
            foreach (int height in new[] {44, 26, 36, 26, 90, 40}) layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.Controls.Add(new Label { Text = "Preview your unsaved rules without creating or sending a message. This does not check mailbox permissions.", Dock = DockStyle.Fill }, 0, 0);
            layout.Controls.Add(new Label { Text = "Original Outlook account", AutoSize = true }, 0, 1);
            testAccount.Items.AddRange(accounts); if (accounts.Length > 0) testAccount.SelectedIndex = 0;
            layout.Controls.Add(testAccount, 0, 2);
            layout.Controls.Add(new Label { Text = "All recipients (To, Cc and Bcc), separated by semicolons", AutoSize = true }, 0, 3);
            layout.Controls.Add(testRecipients, 0, 4);
            var test = new Button { Text = "Test selection", AutoSize = true };
            test.Click += (s, e) => {
                try
                {
                    var config = ReadConfiguration();
                    var recipients = List(testRecipients.Text).Select(Config.Address).ToArray();
                    var decision = RuleEngine.Decide(config, Convert.ToString(testAccount.SelectedItem), recipients, true);
                    testResult.ForeColor = decision.Error == null ? Color.DarkGreen : Color.DarkRed;
                    testResult.Text = !config.Enabled ? "Automatic selection is disabled." : decision.Error != null ? "Send would be blocked: " + decision.Error :
                        decision.Target == null ? "No matching sender. A new draft stays unchanged; an already managed draft restores its original sender." :
                        "Selected sender: " + decision.Target.Id + "\r\nFrom: " + decision.Target.FromAddress + "\r\nAccount: " + decision.Target.Account;
                }
                catch (Exception ex) { testResult.ForeColor = Color.DarkRed; testResult.Text = ex.Message; }
            };
            layout.Controls.Add(test, 0, 5); layout.Controls.Add(testResult, 0, 6);
            tab.Controls.Add(layout); tabs.TabPages.Add(tab);
        }
        public Config ReadConfiguration()
        {
            if (!senders.EndEdit() || !rules.EndEdit()) throw new FormatException("Finish editing the current cell first.");
            var config = new Config { Enabled = enabled.Checked,
                Senders = senders.Rows.Cast<DataGridViewRow>().Select(r => new Sender { Id = Value(r, "id"), Account = Value(r, "account"),
                    Represented = Value(r, "represented"), Alias = Value(r, "alias"), DefaultForAccount = Convert.ToBoolean(r.Cells["default"].Value ?? false) }).ToArray(),
                Rules = rules.Rows.Cast<DataGridViewRow>().Select(r => {
                    int priority;
                    if (!int.TryParse(Value(r, "priority"), out priority)) throw new FormatException("Priority must be a whole number for rule: " + Value(r, "id"));
                    return new Rule { Id = Value(r, "id"), Priority = priority, SenderId = Value(r, "sender"),
                        Recipients = List(Value(r, "recipients")), Domains = List(Value(r, "domains")), Accounts = List(Value(r, "accounts")),
                        IncludeSubdomains = Convert.ToBoolean(r.Cells["subdomains"].Value ?? false) };
                }).ToArray() };
            return SettingsStore.Validate(config);
        }
        void SaveClicked(object sender, EventArgs e)
        {
            try { save(ReadConfiguration()); DialogResult = DialogResult.OK; Close(); }
            catch (Exception ex) { MessageBox.Show(this, "Settings were not saved.\r\n\r\n" + ex.Message, "AutoFrom", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
    }
}
