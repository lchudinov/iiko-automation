# Load Windows.Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell UI Example"
$form.Size = New-Object System.Drawing.Size(300,200)
$form.TopMost = $true  # Set the form to be topmost

# Create a dropdown list
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(50, 20)
$comboBox.Size = New-Object System.Drawing.Size(200, 25)

# Add sample items to the dropdown list
$comboBox.Items.Add("Item 1")
$comboBox.Items.Add("Item 2")
$comboBox.Items.Add("Item 3")
$comboBox.Items.Add("Item 4")

# Create a button
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(50, 70)
$button.Size = New-Object System.Drawing.Size(200, 50)
$button.Text = "Click Me"

# Define the button click event
$button.Add_Click({
    $selectedItem = $comboBox.SelectedItem
    Write-Host "Button Clicked! Selected Item: $selectedItem"
})

# Add controls to the form
$form.Controls.Add($comboBox)
$form.Controls.Add($button)

# Show the form
$form.ShowDialog()
