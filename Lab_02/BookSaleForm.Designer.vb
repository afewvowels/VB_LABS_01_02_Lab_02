<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BookSaleForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.CalculateButton = New System.Windows.Forms.Button()
        Me.ClearButton = New System.Windows.Forms.Button()
        Me.ExitButton = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.PriceTextBox = New System.Windows.Forms.TextBox()
        Me.TitleTextBox = New System.Windows.Forms.TextBox()
        Me.QuantityTextBox = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.DiscountedPriceTextBox = New System.Windows.Forms.TextBox()
        Me.DiscountTextBox = New System.Windows.Forms.TextBox()
        Me.ExtendedPriceTextBox = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.AverageDiscountTextBox = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.DiscountedAmountSumTextBox = New System.Windows.Forms.TextBox()
        Me.DiscountSumTextBox = New System.Windows.Forms.TextBox()
        Me.QuantitySumTextBox = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'CalculateButton
        '
        Me.CalculateButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CalculateButton.Location = New System.Drawing.Point(113, 439)
        Me.CalculateButton.Name = "CalculateButton"
        Me.CalculateButton.Size = New System.Drawing.Size(75, 23)
        Me.CalculateButton.TabIndex = 1
        Me.CalculateButton.Text = "&Calculate"
        Me.CalculateButton.UseVisualStyleBackColor = True
        '
        'ClearButton
        '
        Me.ClearButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ClearButton.Location = New System.Drawing.Point(194, 439)
        Me.ClearButton.Name = "ClearButton"
        Me.ClearButton.Size = New System.Drawing.Size(75, 23)
        Me.ClearButton.TabIndex = 2
        Me.ClearButton.Text = "Clear &Sale"
        Me.ClearButton.UseVisualStyleBackColor = True
        '
        'ExitButton
        '
        Me.ExitButton.Location = New System.Drawing.Point(275, 439)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.Size = New System.Drawing.Size(75, 23)
        Me.ExitButton.TabIndex = 3
        Me.ExitButton.Text = "E&xit"
        Me.ExitButton.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.PriceTextBox)
        Me.GroupBox1.Controls.Add(Me.TitleTextBox)
        Me.GroupBox1.Controls.Add(Me.QuantityTextBox)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(29, 25)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(321, 114)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        '
        'PriceTextBox
        '
        Me.PriceTextBox.Location = New System.Drawing.Point(84, 75)
        Me.PriceTextBox.Name = "PriceTextBox"
        Me.PriceTextBox.Size = New System.Drawing.Size(100, 20)
        Me.PriceTextBox.TabIndex = 5
        '
        'TitleTextBox
        '
        Me.TitleTextBox.Location = New System.Drawing.Point(84, 47)
        Me.TitleTextBox.Name = "TitleTextBox"
        Me.TitleTextBox.Size = New System.Drawing.Size(225, 20)
        Me.TitleTextBox.TabIndex = 4
        '
        'QuantityTextBox
        '
        Me.QuantityTextBox.Location = New System.Drawing.Point(84, 12)
        Me.QuantityTextBox.Name = "QuantityTextBox"
        Me.QuantityTextBox.Size = New System.Drawing.Size(100, 20)
        Me.QuantityTextBox.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 75)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(31, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "&Price"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(27, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "&Title"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "&Quantity"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DiscountedPriceTextBox)
        Me.GroupBox2.Controls.Add(Me.DiscountTextBox)
        Me.GroupBox2.Controls.Add(Me.ExtendedPriceTextBox)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Location = New System.Drawing.Point(29, 145)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(321, 114)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        '
        'DiscountedPriceTextBox
        '
        Me.DiscountedPriceTextBox.Location = New System.Drawing.Point(211, 75)
        Me.DiscountedPriceTextBox.Name = "DiscountedPriceTextBox"
        Me.DiscountedPriceTextBox.ReadOnly = True
        Me.DiscountedPriceTextBox.Size = New System.Drawing.Size(100, 20)
        Me.DiscountedPriceTextBox.TabIndex = 5
        Me.DiscountedPriceTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DiscountTextBox
        '
        Me.DiscountTextBox.Location = New System.Drawing.Point(211, 47)
        Me.DiscountTextBox.Name = "DiscountTextBox"
        Me.DiscountTextBox.ReadOnly = True
        Me.DiscountTextBox.Size = New System.Drawing.Size(100, 20)
        Me.DiscountTextBox.TabIndex = 4
        Me.DiscountTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ExtendedPriceTextBox
        '
        Me.ExtendedPriceTextBox.Location = New System.Drawing.Point(211, 20)
        Me.ExtendedPriceTextBox.Name = "ExtendedPriceTextBox"
        Me.ExtendedPriceTextBox.ReadOnly = True
        Me.ExtendedPriceTextBox.Size = New System.Drawing.Size(100, 20)
        Me.ExtendedPriceTextBox.TabIndex = 3
        Me.ExtendedPriceTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(108, 75)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Discounted Price"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(124, 47)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "15% Discount"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(117, 20)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(79, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Extended Price"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.AverageDiscountTextBox)
        Me.GroupBox3.Controls.Add(Me.Label10)
        Me.GroupBox3.Controls.Add(Me.DiscountedAmountSumTextBox)
        Me.GroupBox3.Controls.Add(Me.DiscountSumTextBox)
        Me.GroupBox3.Controls.Add(Me.QuantitySumTextBox)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Location = New System.Drawing.Point(29, 277)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(321, 140)
        Me.GroupBox3.TabIndex = 7
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Summary"
        '
        'AverageDiscountTextBox
        '
        Me.AverageDiscountTextBox.Location = New System.Drawing.Point(209, 104)
        Me.AverageDiscountTextBox.Name = "AverageDiscountTextBox"
        Me.AverageDiscountTextBox.ReadOnly = True
        Me.AverageDiscountTextBox.Size = New System.Drawing.Size(100, 20)
        Me.AverageDiscountTextBox.TabIndex = 7
        Me.AverageDiscountTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(7, 104)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(92, 13)
        Me.Label10.TabIndex = 6
        Me.Label10.Text = "Average Discount"
        '
        'DiscountedAmountSumTextBox
        '
        Me.DiscountedAmountSumTextBox.Location = New System.Drawing.Point(211, 75)
        Me.DiscountedAmountSumTextBox.Name = "DiscountedAmountSumTextBox"
        Me.DiscountedAmountSumTextBox.ReadOnly = True
        Me.DiscountedAmountSumTextBox.Size = New System.Drawing.Size(100, 20)
        Me.DiscountedAmountSumTextBox.TabIndex = 5
        Me.DiscountedAmountSumTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DiscountSumTextBox
        '
        Me.DiscountSumTextBox.Location = New System.Drawing.Point(211, 47)
        Me.DiscountSumTextBox.Name = "DiscountSumTextBox"
        Me.DiscountSumTextBox.ReadOnly = True
        Me.DiscountSumTextBox.Size = New System.Drawing.Size(100, 20)
        Me.DiscountSumTextBox.TabIndex = 4
        Me.DiscountSumTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'QuantitySumTextBox
        '
        Me.QuantitySumTextBox.Location = New System.Drawing.Point(211, 20)
        Me.QuantitySumTextBox.Name = "QuantitySumTextBox"
        Me.QuantitySumTextBox.ReadOnly = True
        Me.QuantitySumTextBox.Size = New System.Drawing.Size(100, 20)
        Me.QuantitySumTextBox.TabIndex = 3
        Me.QuantitySumTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(7, 75)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(132, 13)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Total Discounted Amounts"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(7, 50)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 13)
        Me.Label8.TabIndex = 1
        Me.Label8.Text = "Total Discounts Given"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(7, 20)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(116, 13)
        Me.Label9.TabIndex = 0
        Me.Label9.Tag = ""
        Me.Label9.Text = "Total Number of Books"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(29, 6)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(321, 23)
        Me.Label11.TabIndex = 8
        Me.Label11.Text = "Book Sales"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'BookSaleForm
        '
        Me.AcceptButton = Me.CalculateButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ClearButton
        Me.ClientSize = New System.Drawing.Size(365, 491)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ExitButton)
        Me.Controls.Add(Me.ClearButton)
        Me.Controls.Add(Me.CalculateButton)
        Me.Name = "BookSaleForm"
        Me.Text = "R 'n R for Reading and Refreshment"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CalculateButton As System.Windows.Forms.Button
    Friend WithEvents ClearButton As System.Windows.Forms.Button
    Friend WithEvents ExitButton As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents PriceTextBox As System.Windows.Forms.TextBox
    Friend WithEvents TitleTextBox As System.Windows.Forms.TextBox
    Friend WithEvents QuantityTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ExtendedPriceTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents DiscountedPriceTextBox As System.Windows.Forms.TextBox
    Friend WithEvents DiscountTextBox As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents AverageDiscountTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents DiscountedAmountSumTextBox As System.Windows.Forms.TextBox
    Friend WithEvents DiscountSumTextBox As System.Windows.Forms.TextBox
    Friend WithEvents QuantitySumTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label

End Class
