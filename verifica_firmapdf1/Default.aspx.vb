
Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        Dim gApp As Acrobat.CAcroApp
        Dim gPDDoc As Acrobat.CAcroPDDoc
        Dim jso As Object
        Dim fname As String
        ' Por ahora hago lo siguiente: porque no se si todos son signature1
        'Con este boton Primero verifico que Signature me trae para luego asignarla abajo en(Dim signatureOne = jso.getField("Signature1")), para poder que funcione

        gApp = CreateObject("AcroExch.App")
        gPDDoc = CreateObject("AcroExch.PDDoc")

        If gPDDoc.Open("C:\temp\maru\vert.pdf") Then
            jso = gPDDoc.GetJSObject

            For i = 0 To jso.numFields - 1
                fname = jso.getNthFieldName(i)
                ListBox1.Items.Add("Campo : " & fname & " valor: " & jso.getField(fname).value)
                TextBox1.Text = jso.getField(fname).value
                Dim signatureOne = jso.getField(TextBox1.Text)

                
                TextBox1.Text = jso.getField(fname).type
                'Extrae la info del firmante, nombre y fecha de la firma

            Next

        End If
    End Sub

    Protected Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
    
        Dim theForm As Acrobat.CAcroPDDoc
        Dim jso As Object
        theForm = CreateObject("AcroExch.PDDoc")
        'Documento de Pdf de Ejemplo en el Diso Local
        theForm.Open("C:\temp\maru\vert5.pdf")
        jso = theForm.GetJSObject

        'Verifica que la firma sea valida

        Dim signatureOne = jso.getField("Signature1")
        Dim oState = signatureOne.SignatureValidate()

        Select Case oState
            Case Is = -1
                ListBox1.Items.Add("Estado : Sin Firma ")
            Case Is = 0
                ListBox1.Items.Add("Estado : Firma en blanco ")
            Case Is = 1
                ListBox1.Items.Add("Estado : No conoce el estado de la firma ")
            Case Is = 2
                ListBox1.Items.Add("Estado : Firma invalida ")
            Case Is = 3
                ListBox1.Items.Add("Estado : La firma es valida, pero la identidad del firmante no se pudo verificar ")
            Case Is = 4
                ListBox1.Items.Add("Estado : Firma e identidad son validas ")

        End Select

        'Extrae la info del firmante, nombre y fecha de la firma
        Dim signatureInformation = signatureOne.signatureInfo

        ListBox1.Items.Add("Firmante " & signatureInformation.name)
        ListBox1.Items.Add("Fecha " & signatureInformation.Date)

        'Extrae la info del certificado
        Dim signatureCertificate = signatureInformation.certificates
        ListBox1.Items.Add("Emitido a : " & signatureCertificate(0).subjectDN.serialNumber)
        ListBox1.Items.Add("Numero de Serie : " & signatureCertificate(0).serialNumber)
        ListBox1.Items.Add("Valido desde : " & signatureCertificate(0).validityStart)
        ListBox1.Items.Add("Valido hasta : " & signatureCertificate(0).validityEnd)
        ListBox1.Items.Add("Para : " & signatureCertificate(0).subjectDN.o)
        ListBox1.Items.Add("Tipo : " & signatureCertificate(0).subjectDN.ou)
        ListBox1.Items.Add("Emitido Por : " & signatureCertificate(0).issuerDN.cn)

    End Sub
End Class
