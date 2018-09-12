' Project: Lab 2
' Author:  Anthony DePinto
' Description: This project inputs sales information for books.  It calculates the extended price and discount for a sale and maintains summary information for all sales.  
'              Uses variables, constants, calculations, error handling and a message box
'
' Date: 10 Sept. 2018
' Student: Keith Smith

Option Explicit On
Option Strict On

Public Class BookSaleForm
    ' Constant variable
    Dim PRICE_DISCOUNT As Decimal = CDec(0.15)
    ' Accumulation variables
    Dim BooksNumberTotal As Integer
    Dim CalculateCounter As Integer
    Dim DiscountSum As Decimal
    Dim DiscountedPriceSum As Decimal
    Dim AverageDiscount As Decimal

    Private Sub ExitButton_Click(sender As System.Object, e As System.EventArgs) Handles ExitButton.Click
        ' exit program
        Me.Close()
    End Sub

    Private Sub ClearButton_Click(sender As System.Object, e As System.EventArgs) Handles ClearButton.Click
        ' Clear previous amounts from the form
        QuantityTextBox.Text = String.Empty
        TitleTextBox.Text = String.Empty
        PriceTextBox.Text = String.Empty

        ExtendedPriceTextBox.Text = String.Empty
        DiscountTextBox.Text = String.Empty
        DiscountedPriceTextBox.Text = String.Empty

        ' Reset cursor position to first data entry field
        QuantityTextBox.Focus()

        ' Can also use these two methods
        ' QuantityTextBox.Text = ""
        ' QuantityTextBox.Clear()
    End Sub

    Private Sub CalculateButton_Click(sender As System.Object, e As System.EventArgs) Handles CalculateButton.Click
        ' Integer variables
        Dim BooksNumberTemp As Integer
        ' Decimal variables
        Dim ItemPrice As Decimal
        Dim ExtendedPrice As Decimal
        Dim Discount As Decimal
        Dim DiscountedPrice As Decimal
        ' String variables
        Dim Title As String
        Dim ErrorString As String

        ' Validate data entry form
        Try
            ' Try to parse values entered in QuantityTextBox
            BooksNumberTemp = Integer.Parse(QuantityTextBox.Text)

            ' Test if TitleTextBox.Text is empty
            If TitleTextBox.Text Is String.Empty Then
                ErrorString = "Title entry field is empty"

                MessageBox.Show(ErrorString,
                                "Title Entry Field Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation)

                ' Reset cursor, testing for empty so no need to highlight label data
                TitleTextBox.Focus()

                ' Not throwing an exception so need to break out of the loop
                Return
            Else
                Title = TitleTextBox.Text
            End If

            ' Try to parse values entered for ItemPrice
            ' and calculate results
            Try
                ItemPrice = Decimal.Parse(PriceTextBox.Text)

                ' If all data is valid, increment calculate counter...
                CalculateCounter += 1

                ' ...and perform calculations
                ExtendedPrice = BooksNumberTemp * ItemPrice
                Discount = ItemPrice * PRICE_DISCOUNT * BooksNumberTemp
                DiscountedPrice = BooksNumberTemp * (ItemPrice - (ItemPrice * PRICE_DISCOUNT))

                ' Calculate results
                ExtendedPriceTextBox.Text = ExtendedPrice.ToString("c")
                DiscountTextBox.Text = Discount.ToString("c")
                DiscountedPriceTextBox.Text = DiscountedPrice.ToString("c")

                BooksNumberTotal += BooksNumberTemp
                DiscountSum += Discount
                DiscountedPriceSum += DiscountedPrice

                QuantitySumTextBox.Text = BooksNumberTotal.ToString()
                DiscountSumTextBox.Text = DiscountSum.ToString("c")
                DiscountedAmountSumTextBox.Text = DiscountedPriceSum.ToString("c")

                AverageDiscount = DiscountedPriceSum / CalculateCounter
                AverageDiscountTextBox.Text = AverageDiscount.ToString("c")

            Catch PriceException As FormatException
                If PriceTextBox.Text Is String.Empty Then
                    ErrorString = "Price entry field is empty."
                Else
                    ErrorString = "Price entry field contains non-integer value."
                End If

                MessageBox.Show(ErrorString,
                                "Price Entry Field Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation)

                ' Alternate means to set multiple properties for an object
                With PriceTextBox
                    .Focus()
                    .SelectAll()
                End With
                ' PriceTextBox.Focus()
                ' PriceTextBox.SelectAll()
            End Try

        Catch QuantityException As FormatException
            If QuantityTextBox.Text Is String.Empty Then
                ErrorString = "Quantity entry field is empty."
            Else
                ErrorString = "Quantity entry field contains a non-integer value."
            End If

            MessageBox.Show(ErrorString,
                            "Quantity Entry Field Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation)

            ' Set QuantityTextBox properties
            With QuantityTextBox
                .Focus()
                .SelectAll()
            End With
            ' QuantityTextBox.Focus() ' move cursor to text box
            ' QuantityTextBox.SelectAll() ' highlight existing data in field
        End Try

    End Sub

    Private Sub QuantityTextBox_TextChanged(sender As Object, e As EventArgs) Handles QuantityTextBox.TextChanged

    End Sub
End Class
