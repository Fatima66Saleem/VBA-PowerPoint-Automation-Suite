VBA GUI Calculator (Built in PowerPoint)
Project Overview
This project is a fully functional Graphical User Interface (GUI) Calculator developed during my 1st Semester. Unlike standard presentations, this project uses VBA (Visual Basic for Applications) to perform real-time mathematical calculations within PowerPoint slides.

Key Features
Basic Arithmetic: Supports Addition, Subtraction, Multiplication, and Division

Event-Driven Programming: Uses CommandButton_Click events to trigger calculations

Clear Functionality: Includes a Clear button to reset all input fields and the result label

Dynamic Input: Users can enter numbers in text boxes and see the result instantly

Technical Logic (VBA Code)
The backend logic is handled via VBA scripts. For example:

Addition: Label4.Caption = Val(TextBox1.Text) + Val(TextBox2.Text)

Reset Logic: Sets TextBox.Text and Label.Caption to empty strings ""

Code Snippets
Addition Operation
Private Sub CommandButton1_Click()
Label4.Caption = Val(TextBox1.Text) + Val(TextBox2.Text)
End Sub

Subtraction Operation
Private Sub CommandButton2_Click()
Label4.Caption = Val(TextBox1.Text) - Val(TextBox2.Text)
End Sub

Multiplication Operation
Private Sub CommandButton3_Click()
Label4.Caption = Val(TextBox1.Text) * Val(TextBox2.Text)
End Sub

Division Operation
Private Sub CommandButton4_Click()
Label4.Caption = Val(TextBox1.Text) / Val(TextBox2.Text)
End Sub

Clear Function
Private Sub CommandButton5_Click()
TextBox1.Text = ""
TextBox2.Text = ""
Label4.Caption = ""
End Sub

How to Use
Open the PowerPoint file

Enable Macros when prompted

Enter numbers in the two text boxes

Click on any operation button (+, -, ×, ÷)

View the result instantly

Click CLEAR to reset

Requirements
Microsoft PowerPoint (2010 or later)
Macros must be enabled

Preview
<img width="972" height="733" alt="Calculator Frontend" src="https://github.com/user-attachments/assets/9320fe33-9709-4266-8332-c9efc14fd677" />
<img width="897" height="536" alt="Calculator code 1" src="https://github.com/user-attachments/assets/12b625b9-5976-4a0f-a4f7-724236fda229" />
<img width="1021" height="610" alt="Calculator code 2" src="https://github.com/user-attachments/assets/ddb9259f-16ca-4885-aadd-ab683cf86292" />


Developed by: Fatima Saleem
Semester: 1st Semester Project
