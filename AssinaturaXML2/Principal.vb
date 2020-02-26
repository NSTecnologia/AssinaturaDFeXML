Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Cryptography.Xml
Imports System.Xml

Public Interface IPrincipal
    Function assinarXML(ByVal conteudo As String, ByVal refURI As String, ByVal cnpjEmitente As String) As String
    Function assinarChaveParaContingencia(chBPe As String, Optional serial As String = "") As String
End Interface

<ClassInterface(ClassInterfaceType.AutoDual)>
Public Class Principal
    Implements IPrincipal
    Public Function assinarXML(ByVal conteudo As String, ByVal refURI As String, ByVal cnpjEmitente As String) As String Implements IPrincipal.assinarXML
        Dim buscaCert As Certificados = New Certificados()
        Dim status, retorno, xmlAssinado As String
        Dim Compartilhados As New Compartilhados

        status = "200"
        retorno = ""
        xmlAssinado = ""

        Compartilhados.gravaLinhaLog("[INICIADO_PROCESSO_DLL]")
        Compartilhados.gravaLinhaLog("RefURI: " & refURI)
        Compartilhados.gravaLinhaLog("CNPJ: " & cnpjEmitente)

        'Busca o certificado com base no XML
        Dim cert As X509Certificate2 = buscaCert.buscaCertificado(cnpjEmitente)

        If cert Is Nothing Then
            status = "-2"
            retorno = "Certificado Digital não encontrado"
        End If

        If (status = "200") Then

            Dim AD As AssinaturaDigital = New AssinaturaDigital()
            Dim resultado As Integer = AD.Assinar(conteudo, refURI, cert)

            status = AD.statusResultado
            retorno = AD.mensagemResultado

            If (status = 0) Then
                xmlAssinado = AD.XMLStringAssinado
            End If

            Compartilhados.gravaLinhaLog("[RESULTADO_ASSINATURA]")
            Compartilhados.gravaLinhaLog("Status: " & status)
            Compartilhados.gravaLinhaLog("MSG Resultado: " & AD.mensagemResultado)
        End If

        Dim json As String
        json = "{"
        json &= """status"":""" & status & ""","
        json &= """msg"":""" & retorno & ""","
        json &= """xmlAssinado"":""" & xmlAssinado.Replace("""", "\""") & """"
        json &= "}"

        Return json

        'Dim parametros As AssinarXmlRet = New AssinarXmlRet With {
        '    .status = status,
        '    .motivo = retorno,
        '    .Xml = xmlAssinado
        '}

        'Dim json As String = JsonConvert.SerializeObject(parametros)

        'gravaLinhaLog("[ASSINAR_XML_RESPOSTA]")
        'gravaLinhaLog(json)

        'Return json
        'Return retorno
    End Function

    Public Function assinarChaveParaContingencia(chBPe As String, Optional serial As String = "") As String
        Dim buscaCert As Certificados = New Certificados()
        Dim status = "200"
        Dim retorno As String = ""
        Dim Compartilhados As New Compartilhados

        Compartilhados.gravaLinhaLog("[ASSINAR_CHAVE_PARA_CONTINGENCIA_DADOS]")
        Compartilhados.gravaLinhaLog("Chave BPe: " & chBPe)
        Compartilhados.gravaLinhaLog("Serial: " & serial)

        'Busca o certificado com base no CNPJ
        Dim cnpj As String = chBPe.Substring(6, 14)
        Dim cert As X509Certificate2 = buscaCert.buscaCertificado(cnpj)

        If cert Is Nothing Then
            status = "-2"
            retorno = "Certificado Digital não encontrado"
        End If

        If (status = "200") Then
            Dim rsaFormatter As New RSACryptoServiceProvider()
            Dim certPK As RSACryptoServiceProvider = DirectCast(cert.PrivateKey, RSACryptoServiceProvider)
            rsaFormatter.ImportParameters(certPK.ExportParameters(True))
            Dim dataToCrypt As Byte() = System.Text.Encoding.UTF8.GetBytes(chBPe)
            Dim sig As Byte() = rsaFormatter.SignData(dataToCrypt, New SHA1CryptoServiceProvider())
            retorno = Convert.ToBase64String(sig)
        End If

        Dim json As String
        json = "{"
        json &= """status"":""" & status & ""","
        json &= """msg"":""" & retorno & """"
        json &= "}"

        Compartilhados.gravaLinhaLog("[ASSINAR_CHAVE_PARA_CONTINGENCIA_RESPOSTA]")
        Compartilhados.gravaLinhaLog(json)

        Return json
    End Function

    Private Function IPrincipal_assinarChaveParaContingencia(chBPe As String, Optional serial As String = "") As String Implements IPrincipal.assinarChaveParaContingencia
        Throw New NotImplementedException()
    End Function
End Class

Class Compartilhados
    Public Sub gravaLinhaLog(ByVal conteudo As String)
        Dim caminho As String = "C:\logDLL\"
        Console.Write(caminho)

        If Not Directory.Exists(caminho) Then
            Directory.CreateDirectory(caminho)
        End If

        Dim data As String = DateTime.Now.ToShortDateString()
        Dim hora As String = DateTime.Now.ToShortTimeString()
        Dim nomeArq As String = data.Replace("/", "")

        Using outputFile As StreamWriter = New StreamWriter(caminho & nomeArq & ".txt", True)
            outputFile.WriteLine(data & " " & hora & " - " & conteudo)
        End Using
    End Sub
End Class

Class Certificados
    Public Sub New()
    End Sub

    Public Function buscaCertificado(ByVal cnpj As String) As X509Certificate2
        Dim Compartilhados As New Compartilhados
        Dim lcerts As X509Certificate2Collection
        Dim lStore As X509Store = New X509Store(StoreName.My, StoreLocation.CurrentUser)

        Compartilhados.gravaLinhaLog("[INICIADO_BUSCA_CERTIFICADO]")

        lStore.Open(OpenFlags.[ReadOnly])
        lcerts = lStore.Certificates
        Dim cert As X509Certificate2 = Nothing

        For Each elemento As X509Certificate2 In lcerts

            If (elemento.Subject.Contains(cnpj)) Or (elemento.SerialNumber.Contains(cnpj)) Then
                Compartilhados.gravaLinhaLog("[CONTEM_CERTIFICADO_CNPJ]")
                cert = elemento
                lStore.Close()
                Return cert
            End If
        Next

        lStore.Close()

        Compartilhados.gravaLinhaLog("[NAO_CONTEM_CERTIFICADO_CNPJ]")

        Return cert
    End Function
End Class


Public Class AssinaturaDigital
    Public Function Assinar(ByVal XMLString As String, ByVal RefUri As String, ByVal X509Cert As X509Certificate2) As Integer
        'Entradas:
        'XMLString: string XML a ser assinada
        'RefUri : Referência da URI a ser assinada (Ex. infNFe
        'X509Cert : certificado digital a ser utilizado na assinatura digital
        '
        'Retornos:
        'Assinar : 0 - Assinatura realizada com sucesso
        '1 - Erro: Problema ao acessar o certificado digital - %exceção%
        '2 - Problemas no certificado digital
        '3 - XML mal formado + exceção
        '4 - A tag de assinatura %RefUri% inexiste
        '5 - A tag de assinatura %RefUri% não é unica
        '6 - Erro Ao assinar o documento - ID deve ser string %RefUri(Atributo)%
        '7 - Erro: Ao assinar o documento - %exceção%
        '
        'XMLStringAssinado : string XML assinada
        '
        'XMLDocAssinado : XMLDocument do XML assinado

        resultado = 0
        msgResultado = "Assinatura realizada com sucesso"

        Dim Compartilhados As New Compartilhados
        Compartilhados.gravaLinhaLog("[INICIADO_ASSINATURA]")

        Try
            'certificado para ser utilizado na assinatura
            Dim _xnome As String = ""

            If X509Cert IsNot Nothing Then
                _xnome = X509Cert.Subject.ToString()
            End If

            Dim _X509Cert As X509Certificate2 = New X509Certificate2()
            Dim store As X509Store = New X509Store("MY", StoreLocation.CurrentUser)
            store.Open(OpenFlags.[ReadOnly] Or OpenFlags.OpenExistingOnly)
            Dim collection As X509Certificate2Collection = CType(store.Certificates, X509Certificate2Collection)
            Dim collection1 As X509Certificate2Collection = CType(collection.Find(X509FindType.FindBySubjectDistinguishedName, _xnome, False), X509Certificate2Collection)

            If collection1.Count = 0 Then
                resultado = 2
                msgResultado = "Problemas no certificado digital"
            Else
                'certificado ok
                _X509Cert = collection1(0)
                Dim x As String
                x = _X509Cert.GetKeyAlgorithm().ToString()
                'cria um novo documento XML
                Dim doc As XmlDocument = New XmlDocument()
                'formata o documento para ignorar os espaços em branco
                doc.PreserveWhitespace = False

                'lê o arquivo XML passado utilizando o nome dele
                Try
                    doc.LoadXml(XMLString)
                    'Verifica se a tag a ser assinada existe é única
                    Dim qtdeRefUri As Integer = doc.GetElementsByTagName(RefUri).Count

                    If qtdeRefUri = 0 Then
                        'a URI indicada não existe
                        resultado = 4
                        msgResultado = "A tag de assinatura " & RefUri.Trim() & " inexiste"
                        'Existe mais de uma tag a ser assinada
                    Else

                        If qtdeRefUri > 1 Then
                            'existe mais de uma URI indicada
                            resultado = 5
                            msgResultado = "A tag de assinatura " & RefUri.Trim() & " não é unica"
                        Else

                            Try
                                'Cria um objeto SignedXml
                                Dim signedXml As SignedXml = New SignedXml(doc)

                                'Adiciona a chave ao documento SignedXml
                                signedXml.SigningKey = _X509Cert.PrivateKey

                                'Cria uma referência a ser assinada
                                Dim reference As Reference = New Reference()

                                'Pega a URI que deve ser assinada
                                Dim _Uri As XmlAttributeCollection = doc.GetElementsByTagName(RefUri).Item(0).Attributes

                                For Each _atributo As XmlAttribute In _Uri

                                    If _atributo.Name = "Id" Then
                                        reference.Uri = "#" & _atributo.InnerText
                                    End If
                                Next

                                Dim env As XmlDsigEnvelopedSignatureTransform = New XmlDsigEnvelopedSignatureTransform()
                                reference.AddTransform(env)
                                Dim c14 As XmlDsigC14NTransform = New XmlDsigC14NTransform()
                                reference.AddTransform(c14)
                                signedXml.AddReference(reference)
                                Dim keyInfo As KeyInfo = New KeyInfo()
                                keyInfo.AddClause(New KeyInfoX509Data(_X509Cert))
                                signedXml.KeyInfo = keyInfo
                                signedXml.ComputeSignature()
                                Dim xmlDigitalSignature As XmlElement = signedXml.GetXml()
                                doc.DocumentElement.AppendChild(doc.ImportNode(xmlDigitalSignature, True))
                                XMLDoc = New XmlDocument()
                                XMLDoc.PreserveWhitespace = False
                                XMLDoc = doc
                            Catch caught As Exception
                                resultado = 7
                                msgResultado = "Erro: Ao assinar o documento - " & caught.Message
                            End Try
                        End If
                    End If

                Catch caught As Exception
                    resultado = 3
                    msgResultado = "Erro: XML mal formado - " & caught.Message
                End Try
            End If

        Catch caught As Exception
            resultado = 1
            msgResultado = "Erro: Problema ao acessar o certificado digital" & caught.Message
        End Try

        Return resultado
    End Function

    Private msgResultado As String
    Private resultado As Integer
    Private XMLDoc As XmlDocument

    Public ReadOnly Property XMLDocAssinado As XmlDocument
        Get
            Return XMLDoc
        End Get
    End Property

    Public ReadOnly Property XMLStringAssinado As String
        Get
            Return XMLDoc.OuterXml
        End Get
    End Property

    Public ReadOnly Property statusResultado As String
        Get
            Return resultado
        End Get
    End Property

    Public ReadOnly Property mensagemResultado As String
        Get
            Return msgResultado
        End Get
    End Property
End Class