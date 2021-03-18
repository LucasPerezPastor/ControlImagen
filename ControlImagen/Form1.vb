Imports System.IO

Public Structure Imagen_Base
    Public Direccion As String
    Public Texto As String
End Structure

Public Structure Imagen_Fusion
    Public Imagen1 As Imagen_Base
    Public Imagen2 As Imagen_Base
    Public FusionImagen As Imagen_Base
End Structure

Public Class Form1
    Dim Datos_Programa As Imagen_Fusion
    Dim ContadorRegresivo As Integer
    Dim Imagen_Izda, Imagen_Drcha, Imagen_Central As Image
    Dim Pulsado As Color = Color.GreenYellow

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Datos_Programa.Imagen1.Direccion = ""
        If Imagen_Izda Is Nothing Then
        Else
            Imagen_Izda.Dispose()
        End If
        Carga_Imagen(Datos_Programa.Imagen1.Direccion, Imagen_Izda, Label1.Text)
        Label1.Visible = True
        PictureBox1.Image = Imagen_Izda
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        Button2.BackColor = SystemColors.Control
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Imagen_Izda Is Nothing Then
        Else
            Imagen_Izda.RotateFlip(RotateFlipType.RotateNoneFlipX)
            ' Este método refleja la imagen en horizontal, pero se puede reflejar en vertical (RotateFlipType.RotateNoneFlipY) o en horizontal y vertical a la vez (RotateFlipType.RotateNoneFlipXY). 
            PictureBox1.Image = Imagen_Izda
            If Button2.BackColor = SystemColors.Control Then
                Button2.BackColor = Pulsado
            Else
                Button2.BackColor = SystemColors.Control
            End If
        End If

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Datos_Programa.Imagen2.Direccion = ""
        If Imagen_Drcha Is Nothing Then
        Else
            Imagen_Drcha.Dispose()
        End If
        Carga_Imagen(Datos_Programa.Imagen2.Direccion, Imagen_Drcha, Label2.Text)
        Label2.Visible = True
        PictureBox2.Image = Imagen_Drcha
        PictureBox2.SizeMode = PictureBoxSizeMode.Zoom
        Button5.BackColor = SystemColors.Control
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If Imagen_Drcha Is Nothing Then
        Else
            Imagen_Drcha.RotateFlip(RotateFlipType.RotateNoneFlipX)
            ' Este método refleja la imagen en horizontal, pero se puede reflejar en vertical (RotateFlipType.RotateNoneFlipY) o en horizontal y vertical a la vez (RotateFlipType.RotateNoneFlipXY). 
            PictureBox2.Image = Imagen_Drcha
            If Button5.BackColor = SystemColors.Control Then
                Button5.BackColor = Pulsado
            Else
                Button5.BackColor = SystemColors.Control
            End If
        End If



    End Sub

    Private Sub Fusion_Imagenes()
        Dim Ancho, Alto As Integer
        Dim Fuente As System.Drawing.Font = New Font("Arial", 10, FontStyle.Underline)
        Dim Pincel As System.Drawing.Brush = New SolidBrush(Color.Black)
        Dim TextoRef As String
        Dim valorFuente As Single
        Dim Longitud, offset As Integer
        Dim TamFuente As SizeF
        Dim Relleno As Integer

        valorFuente = 10
        Ancho = Math.Max(Imagen_Izda.Width, Imagen_Drcha.Width)
        Alto = Imagen_Izda.Height + Imagen_Drcha.Height

        Longitud = Math.Max(TextBox1.Text.Length, TextBox2.Text.Length)
        Fuente = New Font("Arial", valorFuente, FontStyle.Underline)
        If TextBox1.Text.Length > TextBox2.Text.Length Then
            TextoRef = TextBox1.Text
        Else
            TextoRef = TextBox2.Text
        End If
        offset = 0
        Relleno = TextoRef.Length / 2
        TextoRef = TextoRef.PadLeft(Relleno + TextoRef.Length, "*")
        TextoRef = TextoRef.PadRight(Relleno + TextoRef.Length, "*")
        Try
            If Imagen_Central Is Nothing Then
            Else
                Imagen_Central.Dispose()
            End If

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        Imagen_Central = New Bitmap(Ancho, Alto)

        Dim gr As Graphics = Graphics.FromImage(Imagen_Central)

        If Longitud > 0 Then
            Do
                TamFuente = gr.MeasureString(TextoRef, Fuente)
                If Ancho > TamFuente.Width Then
                    valorFuente += 1
                End If
                Fuente.Dispose()
                Fuente = New Font("Arial", valorFuente, FontStyle.Underline)
            Loop Until Ancho < TamFuente.Width
            valorFuente -= 5
            Fuente.Dispose()
            Fuente = New Font("Arial", valorFuente, FontStyle.Underline)

            valorFuente = TamFuente.Height
            If TextBox1.Text <> "" Then Alto = Alto + valorFuente * 2
            If TextBox2.Text <> "" Then Alto = Alto + valorFuente * 2
            Imagen_Central.Dispose()
            Imagen_Central = New Bitmap(Ancho, Alto)
        End If
        gr = Graphics.FromImage(Imagen_Central)
        gr.Clear(Color.White)
        If TextBox1.Text <> "" Then
            TamFuente = gr.MeasureString(TextBox1.Text, Fuente)
            offset = (Ancho - TamFuente.Width) / 2
            gr.DrawString(TextBox1.Text, Fuente, Pincel, offset, 0)
        End If

        gr.DrawImage(Imagen_Izda, 0, valorFuente * 2, Imagen_Izda.Width, Imagen_Izda.Height)
        If TextBox2.Text <> "" Then
            TamFuente = gr.MeasureString(TextBox2.Text, Fuente)
            offset = (Ancho - TamFuente.Width) / 2
            gr.DrawString(TextBox2.Text, Fuente, Pincel, offset, Imagen_Izda.Height + valorFuente * 2)
        End If


        gr.DrawImage(Imagen_Drcha, 0, (Imagen_Izda.Height + (valorFuente * 4)), Imagen_Drcha.Width, Imagen_Drcha.Height)
        PictureBox3.Image = Imagen_Central
        PictureBox3.SizeMode = PictureBoxSizeMode.Zoom
        Label3.Text = "Tamaño imagen:" + Imagen_Central.Width.ToString + "x" + Imagen_Central.Height.ToString
        Label3.Visible = True






    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Fusion_Imagenes()
    End Sub


    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim SaveFileDialog As New SaveFileDialog
        Dim Direccion_Guardado As String
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.ToString
        SaveFileDialog.Filter = "Archivos Imagen (*.jpg)|*.jpg"
        ' SaveFileDialog.Multiselect = True
        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Try
                Direccion_Guardado = SaveFileDialog.FileName.ToString
                Imagen_Central.Save(Direccion_Guardado, Drawing.Imaging.ImageFormat.Jpeg)
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Carga_Imagen(ByRef Direccion As String, ByRef Imagen As Image, ByRef Etiqueta As String)
        Dim OpenFileDialog As New OpenFileDialog

        If Direccion = "" Then
            OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.ToString
            OpenFileDialog.Filter = "Archivos Imagen (*.jpg)|*.jpg|Todos los archivos (*.*)|*.*"
            OpenFileDialog.Multiselect = True
            If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
                Try
                    Direccion = OpenFileDialog.FileName.ToString
                    Imagen = System.Drawing.Image.FromFile(Direccion)
                    'If Imagen.Width < Imagen.Height Then
                    'Imagen.RotateFlip(RotateFlipType.Rotate270FlipNone)
                    'End If
                    Etiqueta = Imagen.Width.ToString + "x" + Imagen.Height.ToString
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
            End If
        Else
            Try
                If System.IO.File.Exists(Direccion) Then
                    Imagen = System.Drawing.Image.FromFile(Direccion)
                    Etiqueta = Imagen.Width.ToString + "x" + Imagen.Height.ToString
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try



        End If

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Comprobar si hay más de un parámetro,
        ' el parámetro CERO es el nombre del ejecutable

        Dim Archivo As Boolean

        Dim Direccion_Archivo As String
        Archivo = False

        Direccion_Archivo = ""

        Datos_Programa.Imagen1.Direccion = ""
        Datos_Programa.Imagen1.Texto = ""
        Datos_Programa.Imagen2.Direccion = ""
        Datos_Programa.Imagen2.Texto = ""
        Datos_Programa.FusionImagen.Direccion = ""
        Datos_Programa.FusionImagen.Texto = ""


        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                'Si recibimos solo un parámetro indica un archivo donde dentro vienen 
                'incluidas las direcciones de las imágenes a tratar , así como el archivo donde guardarlas.
                'Este archivo debe incluir 5 líneas de texto.
                ' Direccion y nombre de la Imagen1
                ' Texto de la Imagen1
                ' Direccion y nombre de la Imagen2
                ' Texto de la Imagen2
                ' Direccion y nombre del archibo a generar

                Direccion_Archivo = Environment.GetCommandLineArgs(1)
                Archivo = True
            Else
                If Environment.GetCommandLineArgs.Length > 2 Then
                    ' Si recibimos más de u parámetro 
                    ' El primer parámetro sera la dirección y nombre de la primera imagen
                    ' El segundo parámetro será la dirección y nombre de la segunda imagen
                    ' El tercer parámetro será la dirección y nombre del archivo a generar
                    Datos_Programa.Imagen1.Direccion = Environment.GetCommandLineArgs(1)
                    Datos_Programa.Imagen2.Direccion = Environment.GetCommandLineArgs(2)
                    Datos_Programa.FusionImagen.Direccion = Environment.GetCommandLineArgs(3)
                    Archivo = False

                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        If Archivo Then

            If Direccion_Archivo <> "" Then
                Dim sLine As String = ""
                Dim arrText As New ArrayList()
                Try
                    If System.IO.File.Exists(Direccion_Archivo) Then
                        Dim objReader As New StreamReader(Direccion_Archivo)
                        Do
                            sLine = objReader.ReadLine()
                            If Not sLine Is Nothing Then
                                arrText.Add(sLine)
                            End If
                        Loop Until sLine Is Nothing
                        objReader.Close()
                        objReader.Dispose()
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                If arrText.Count = 5 Then
                    Datos_Programa.Imagen1.Direccion = arrText(0)
                    Datos_Programa.Imagen1.Texto = arrText(1)
                    Datos_Programa.Imagen2.Direccion = arrText(2)
                    Datos_Programa.Imagen2.Texto = arrText(3)
                    Datos_Programa.FusionImagen.Direccion = arrText(4)
                Else
                    Console.WriteLine("Datos insuficientes en el archivo: " + Direccion_Archivo)

                End If

            End If
        End If



        If Datos_Programa.Imagen1.Direccion <> "" Then
            TextBox1.Text = Datos_Programa.Imagen1.Texto
            Carga_Imagen(Datos_Programa.Imagen1.Direccion, Imagen_Izda, Label1.Text)
            Label1.Visible = True
            ' Colocamnos la imagen en formato horizontal
            If Imagen_Izda.Height > Imagen_Izda.Width Then Imagen_Izda.RotateFlip(RotateFlipType.Rotate270FlipNone)
            PictureBox1.Image = Imagen_Izda
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        End If
        If Datos_Programa.Imagen2.Direccion <> "" Then
            TextBox2.Text = Datos_Programa.Imagen2.Texto
            Carga_Imagen(Datos_Programa.Imagen2.Direccion, Imagen_Drcha, Label2.Text)
            Label2.Visible = True
            ' PictureBox2.Image = Imagen_Drcha
            PictureBox2.SizeMode = PictureBoxSizeMode.Zoom
            ' La imagen es colocada en efecto espejo de forma automática
            Imagen_Drcha.RotateFlip(RotateFlipType.RotateNoneFlipX)
            ' Este método refleja la imagen en horizontal, pero se puede reflejar en vertical (RotateFlipType.RotateNoneFlipY) o en horizontal y vertical a la vez (RotateFlipType.RotateNoneFlipXY). 
            ' Colocamnos la imagen en formato horizontal
            If Imagen_Drcha.Height > Imagen_Drcha.Width Then Imagen_Drcha.RotateFlip(RotateFlipType.Rotate270FlipNone)
            PictureBox2.Image = Imagen_Drcha
            Button5.BackColor = Pulsado
        End If
        If ((Datos_Programa.Imagen1.Direccion <> "") And (Datos_Programa.Imagen2.Direccion <> "")) Then
            Fusion_Imagenes()
            If Datos_Programa.FusionImagen.Direccion <> "" Then
                Imagen_Central.Save(Datos_Programa.FusionImagen.Direccion, Drawing.Imaging.ImageFormat.Jpeg)
            End If
            'Iniciamos la cuneta atrás para el cierre automático
            ContadorRegresivo = 10
            Button8.Visible = True
            Button8.Text = ContadorRegresivo
            Timer1.Start()
        End If





    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Imagen_Izda Is Nothing Then
        Else
            Imagen_Izda.RotateFlip(RotateFlipType.Rotate270FlipNone)
            ' Este método refleja la imagen en horizontal, pero se puede reflejar en vertical (RotateFlipType.RotateNoneFlipY) o en horizontal y vertical a la vez (RotateFlipType.RotateNoneFlipXY). 
            PictureBox1.Image = Imagen_Izda

        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Imagen_Drcha Is Nothing Then
        Else
            Imagen_Drcha.RotateFlip(RotateFlipType.Rotate270FlipNone)
            ' Este método refleja la imagen en horizontal, pero se puede reflejar en vertical (RotateFlipType.RotateNoneFlipY) o en horizontal y vertical a la vez (RotateFlipType.RotateNoneFlipXY). 
            PictureBox2.Image = Imagen_Drcha
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ContadorRegresivo > 0 Then
            ' Display the new time left
            ' by updating the Time Left label.
            ContadorRegresivo -= 1
            Button8.Text = ContadorRegresivo
        Else
            ' If the user ran out of time, stop the timer, show
            ' a MessageBox, and fill in the answers.
            Timer1.Stop()
            ' Si la cuenta atrás a llegado a cero cerramos el programa
            Me.Close()
        End If

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        ' Paramos la cuenta atrás 
        If Timer1.Enabled Then Timer1.Stop()
        Button8.Visible = False

    End Sub
End Class
