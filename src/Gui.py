
import clr
from BeamsOpt import BeamOptimzer
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import (
    Application,
    Form,
    Label,
    TextBox,
    Button,
    CheckBox,
    GroupBox,
    ProgressBar,
    TrackBar,
    Panel,
    TableLayoutPanel,
    ColumnStyle,
    RowStyle,
    SizeType,
    DockStyle,
    FormStartPosition,
    Padding,
    BorderStyle,
    FlatStyle,
    Cursors,
    AnchorStyles,
)
from System.Drawing import Size, Point, Color, ContentAlignment


class BeamForm(Form):

    def __init__(self):
        super().__init__()
        self.beam_opt = BeamOptimzer()

        # --------------------------------------------------
        # Window
        # --------------------------------------------------
        self.Text = "Beam Design | Parameters"
        self.Size = Size(620, 720)
        self.MinimumSize = Size(480, 620)
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.FromArgb(245, 247, 250)

        # --------------------------------------------------
        # Colors
        # --------------------------------------------------
        self.bg = Color.FromArgb(245, 247, 250)
        self.white = Color.White
        self.text_color = Color.FromArgb(38, 50, 65)
        self.muted = Color.FromArgb(105, 117, 132)
        self.accent = Color.FromArgb(36, 99, 170)

        # --------------------------------------------------
        # Main layout
        # --------------------------------------------------
        self.layout = TableLayoutPanel()
        self.layout.Dock = DockStyle.Fill
        self.layout.Padding = Padding(18, 12, 18, 12)
        self.layout.BackColor = self.bg
        self.layout.ColumnCount = 1
        self.layout.RowCount = 7

        self.layout.ColumnStyles.Add(
            ColumnStyle(SizeType.Percent, 100)
        )

        row_heights = [12, 27, 19, 14, 9, 14, 5]

        for height in row_heights:
            self.layout.RowStyles.Add(
                RowStyle(SizeType.Percent, height)
            )

        self.Controls.Add(self.layout)

        # --------------------------------------------------
        # Header
        # --------------------------------------------------
        self.header = Panel()
        self.header.Dock = DockStyle.Fill
        self.header.BackColor = self.bg

        self.title = Label()
        self.title.Text = "BEAM DESIGN"
        self.title.Location = Point(2, 2)
        self.title.AutoSize = True
        self.title.ForeColor = self.text_color

        self.subtitle = Label()
        self.subtitle.Text = (
            "Reinforcement parameters and section constraints"
        )
        self.subtitle.Location = Point(4, 28)
        self.subtitle.AutoSize = True
        self.subtitle.ForeColor = self.muted

        self.header.Controls.Add(self.title)
        self.header.Controls.Add(self.subtitle)
        self.layout.Controls.Add(self.header, 0, 0)

        # --------------------------------------------------
        # 1. Steel ratio
        # --------------------------------------------------
        self.steel_group = self.make_group(
            "01   LONGITUDINAL STEEL RATIO"
        )
        self.layout.Controls.Add(self.steel_group, 0, 1)

        self.ratio_label = Label()
        self.ratio_label.Text = "1.20 %"
        self.ratio_label.Location = Point(18, 28)
        self.ratio_label.AutoSize = True
        self.ratio_label.ForeColor = self.accent

        self.ratio_range = Label()
        self.ratio_range.Text = (
            "MIN  0.20 %                         MAX  2.00 %"
        )
        self.ratio_range.Location = Point(22, 59)
        self.ratio_range.AutoSize = True
        self.ratio_range.ForeColor = self.muted

        self.ratio_track = TrackBar()
        self.ratio_track.Minimum = 20
        self.ratio_track.Maximum = 200
        self.ratio_track.Value = 120
        self.ratio_track.TickFrequency = 20
        self.ratio_track.SmallChange = 1
        self.ratio_track.LargeChange = 10
        self.ratio_track.Location = Point(12, 76)
        self.ratio_track.Size = Size(480, 42)
        self.ratio_track.Anchor = (
            AnchorStyles.Top
            | AnchorStyles.Left
            | AnchorStyles.Right
        )
        self.ratio_track.ValueChanged += self.ratio_changed

        self.ratio_info = Label()
        self.ratio_info.Location = Point(22, 119)
        self.ratio_info.Size = Size(480, 22)
        self.ratio_info.ForeColor = self.muted
        self.ratio_info.Anchor = (
            AnchorStyles.Top
            | AnchorStyles.Left
            | AnchorStyles.Right
        )

        self.steel_group.Controls.Add(self.ratio_label)
        self.steel_group.Controls.Add(self.ratio_range)
        self.steel_group.Controls.Add(self.ratio_track)
        self.steel_group.Controls.Add(self.ratio_info)

        # --------------------------------------------------
        # 2. Section constraints
        # --------------------------------------------------
        self.width_group = self.make_group(
            "02   SECTION CONSTRAINTS"
        )
        self.layout.Controls.Add(self.width_group, 0, 2)

        self.width_enabled = CheckBox()
        self.width_enabled.Text = "Apply minimum beam height"
        self.width_enabled.Location = Point(18, 32)
        self.width_enabled.AutoSize = True
        self.width_enabled.ForeColor = self.text_color
        self.width_enabled.Checked = False
        self.width_enabled.CheckedChanged += self.toggle_width

        self.width_label = Label()
        self.width_label.Text = "b min (m)"
        self.width_label.Location = Point(22, 70)
        self.width_label.AutoSize = True
        self.width_label.ForeColor = self.muted

        self.width_input = TextBox()
        self.width_input.Text = "0.21"
        self.width_input.Location = Point(125, 66)
        self.width_input.Size = Size(105, 25)
        self.width_input.Enabled = False
        self.width_input.BorderStyle = BorderStyle.FixedSingle

        self.width_unit = Label()
        self.width_unit.Text = "metres"
        self.width_unit.Location = Point(239, 70)
        self.width_unit.AutoSize = True
        self.width_unit.ForeColor = self.muted

        self.width_group.Controls.Add(self.width_enabled)
        self.width_group.Controls.Add(self.width_label)
        self.width_group.Controls.Add(self.width_input)
        self.width_group.Controls.Add(self.width_unit)

        # --------------------------------------------------
        # 3. Optimization
        # --------------------------------------------------
        self.options_group = self.make_group(
            "03   OPTIMIZATION"
        )
        self.layout.Controls.Add(self.options_group, 0, 3)

        self.reduce_high = CheckBox()
        self.reduce_high.Text = (
            "Reduce Height where possible"
        )
        self.reduce_high.Location = Point(18, 32)
        self.reduce_high.AutoSize = True
        self.reduce_high.ForeColor = self.text_color
        self.reduce_high.Checked = False

        self.options_group.Controls.Add(self.reduce_high)

        # --------------------------------------------------
        # 4. Calculate button
        # --------------------------------------------------
        self.calculate_button = Button()
        self.calculate_button.Text = "Calculate"
        self.calculate_button.Dock = DockStyle.Fill
        self.calculate_button.Margin = Padding(0, 6, 0, 6)
        self.calculate_button.FlatStyle = FlatStyle.Flat
        self.calculate_button.FlatAppearance.BorderSize = 0
        self.calculate_button.BackColor = self.accent
        self.calculate_button.ForeColor = Color.White
        self.calculate_button.Cursor = Cursors.Hand
        self.calculate_button.Click += self.calculate

        self.layout.Controls.Add(self.calculate_button, 0, 4)

        # --------------------------------------------------
        # 5. Calculation result
        # --------------------------------------------------
        self.output_group = self.make_group(
            "CALCULATION RESULT"
        )
        self.layout.Controls.Add(self.output_group, 0, 5)

        self.output_label = Label()
        self.output_label.Location = Point(18, 32)
        self.output_label.AutoSize = True
        self.output_label.ForeColor = self.accent

        # Progress bar
        self.output_bar = ProgressBar()
        self.output_bar.Minimum = 0
        self.output_bar.Maximum = 100
        self.output_bar.Value = 0
        self.output_bar.Location = Point(20, 65)
        self.output_bar.Size = Size(420, 20)
        self.output_bar.Anchor = (
            AnchorStyles.Top
            | AnchorStyles.Left
            | AnchorStyles.Right
        )
        self.output_bar.Visible = False

        # Percentage label
        self.progress_percent = Label()
        self.progress_percent.Text = "0%"
        self.progress_percent.Location = Point(450, 65)
        self.progress_percent.Size = Size(50, 20)
        self.progress_percent.AutoSize = False
        self.progress_percent.TextAlign = ContentAlignment.MiddleLeft
        self.progress_percent.ForeColor = self.accent
        self.progress_percent.Visible = False

        self.output_group.Controls.Add(self.output_label)
        self.output_group.Controls.Add(self.output_bar)
        self.output_group.Controls.Add(self.progress_percent)

        # --------------------------------------------------
        # 6. Footer
        # --------------------------------------------------
        self.footer = Label()
        self.footer.Text = "Structural design parameters"
        self.footer.Dock = DockStyle.Fill
        self.footer.ForeColor = self.muted
        self.footer.TextAlign = ContentAlignment.MiddleCenter

        self.layout.Controls.Add(self.footer, 0, 6)

        self.update_ratio_info()

        self.Shown += self.form_shown

    def form_shown(self, sender, event):
        self.calculate_button.Enabled = False
        self.calculate_button.Visible =  False
        self.output_bar.Visible =  True
        self.output_label.Text = "Loading beams ..."
        load_counter = 0 
        for progress_percent in self.beam_opt.get_current_selection() : 
            self.update_progress(progress_percent)
            load_counter = load_counter + 1 
            self.output_label.Text = f"{load_counter} Beam(s) Loaded ....."
        # Enable the button
        self.calculate_button.Enabled = True
        self.calculate_button.Visible =  True
        self.output_bar.Visible =  False





    # ------------------------------------------------------
    # Group helper
    # ------------------------------------------------------
    def make_group(self, title):
        group = GroupBox()
        group.Text = title
        group.Dock = DockStyle.Fill
        group.Margin = Padding(0, 5, 0, 5)
        group.BackColor = self.white
        group.ForeColor = self.text_color
        group.Padding = Padding(10)
        return group

    # ------------------------------------------------------
    # Steel ratio
    # ------------------------------------------------------
    @property
    def steel_ratio(self):
        return self.ratio_track.Value / 100.0

    def ratio_changed(self, sender, event):
        self.ratio_label.Text = f"{self.steel_ratio:.2f} %"
        self.update_ratio_info()

    def update_ratio_info(self):
        ratio = self.steel_ratio
        self.beam_opt.set_ratio_steel(ratio)

        if ratio <= 0.2:
            description = (
                "Low reinforcement - verify code minimum"
            )
        elif ratio <= 0.5:
            description = "Lightly reinforced"
        elif ratio <= 1.0:
            description = "Common practical range"
        elif ratio <= 1.5:
            description = "Moderately to heavily reinforced"
        else:
            description = "Heavily reinforced - check detailing"

        self.ratio_info.Text = description

    # ------------------------------------------------------
    # Progress display
    # ------------------------------------------------------
    def update_progress(self, percent):
        percent = max(0, min(100, int(percent)))

        self.output_bar.Visible = True
        self.progress_percent.Visible = True

        self.output_bar.Value = percent
        self.progress_percent.Text = f"{percent}%"

        Application.DoEvents()

    # ------------------------------------------------------
    # Set calculated ratio
    # Example: form.set_steel_ratio(0.85)
    # ------------------------------------------------------
    def set_steel_ratio(self, ratio):
        ratio = float(ratio)

        if not 0.2 <= ratio <= 2.0:
            raise ValueError(
                "Steel ratio must be between 0.2% and 2.0%."
            )



        self.ratio_track.Value = int(round(ratio * 100))
        self.ratio_label.Text = f"{self.steel_ratio:.2f} %"
        self.update_ratio_info()

    # ------------------------------------------------------
    # Minimum width option
    # ------------------------------------------------------
    def toggle_width(self, sender, event):
        self.width_input.Enabled = self.width_enabled.Checked

    # ------------------------------------------------------
    # Calculate
    # ------------------------------------------------------
    def calculate(self, sender, event):
        self.output_bar.Visible =  True
        self.calculate_button.Enabled = False
        self.calculate_button.Visible =  False

        self.output_label.Text = "Calculation started..."

        self.update_progress(0)

        try:
            input_ratio = self.steel_ratio

            # Validate minimum beam width if enabled.
            if self.width_enabled.Checked:
                try:
                    b_min = float(
                        self.width_input.Text.replace(",", ".")
                    )
                    if b_min <= 0:
                        raise ValueError
                except ValueError:
                    self.width_input.Focus()
                    self.width_input.SelectAll()
                    self.output_label.Text = (
                        "Enter a valid minimum width."
                    )
                    self.output_bar.Visible = False
                    self.progress_percent.Visible = False
                    return
            else:
                b_min = None

            reduce_high = self.reduce_high.Checked

            # --------------------------------------------------
            # PLACEHOLDER CALCULATION
            # Replace this section with the actual beam design.
            # The percentages below represent illustrative stages.
            # --------------------------------------------------





            calculated_ratio = input_ratio
            self.beam_opt.set_ratio_steel(calculated_ratio)

            self.output_label.Text = "Loading selected Beams..."


            # Update result.
            self.set_steel_ratio(calculated_ratio)
            optimzed_counter = 0
            self.update_progress(0)
            self.output_label.Text = f"{optimzed_counter} Beam(s) Optimzed ....."
            for progress_percent in self.beam_opt.optimizeSelectedBeams(h_min = b_min ,reduce_h = reduce_high) : 
                self.update_progress(progress_percent)
                optimzed_counter = optimzed_counter + 1 
                self.output_label.Text = f"{optimzed_counter} Beam(s) Optimzed ....."

            self.update_progress(100)
            self.output_label.Text += "\nCalculation completed."
            self.output_bar.Visible =  False



            

            print("Input steel ratio:", input_ratio, "%")
            print("Calculated steel ratio:", calculated_ratio, "%")
            print("Minimum beam width:", b_min, "m")
            print("Reduce high:", reduce_high)



        except Exception as exc:
            self.output_label.Text = f"Calculation failed: {exc}"
            raise

        finally:
            self.calculate_button.Enabled = True
        self.Close()

    # ------------------------------------------------------
    # Close
    # ------------------------------------------------------
    def close_form(self, sender, event):
        self.Close()


import traceback

if __name__ == "__main__":
    try:
        Application.EnableVisualStyles()
        form = BeamForm()
        Application.Run(form)

    except Exception:
        traceback.print_exc()

