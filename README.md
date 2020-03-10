# AssinaturaXML2
Faz a assinatura de um XML utilizando certificado digital usando uma dll

----------

## Primeiros passos:

### Gerando um .tlb da dll:

1. Pegue a AssinaturaXML.dll e copie para a pasta:
	- C:\Windows\System32 (caso seu Windows for x86);
	- C:\Windows\SysWOW64 (caso seu Windows for x64);
2. Abra o Windows PowerShell como Administrador e utilize este comando: cd C:\Windows\Microsoft.NET\Framework\v4.0.30319
3. Execute no Windows PowerShell o seguinte comando:
	- Caso x86: .\RegAsm.exe C:\Windows\System32\AssinaturaXML.dll /codebase /tlb:C:\Windows\System32\AssinaturaXML.tlb
	- Caso x64: .\RegAsm.exe C:\Windows\SysWOW64\AssinaturaXML.dll /codebase /tlb:C:\Windows\SysWOW64\AssinaturaXML.tlb
	
	**ATENÇÃO:** caso não tenha sido criado o arquivo .tlb voce deve executar novamente o passo 3
	
4. Após a criação do arquivo .tlb voce deve referencia-lo no seu projeto VBA para que seja possivel utilizar suas funcionalidades

----------

## Utilizando o AssinaturaXML:


### Realizando a refencia no Projeto em VB6:

1. Abra seu projeto VBA e vá ate a aba 'Projeto':

![Screenshot_1](https://user-images.githubusercontent.com/54732019/75123207-339fbd80-5684-11ea-9d79-63d6abe2df59.png)

2. Dentro de 'Projeto' vá até Referências:

![Screenshot_2](https://user-images.githubusercontent.com/54732019/75123213-49ad7e00-5684-11ea-85e8-d803712a1045.png)

3. No formulario de referencias busque pelo arquivo .tlb gerado anteriormente:

![3](https://user-images.githubusercontent.com/54732019/76355330-ddbd4d80-62f2-11ea-926c-b0e859c4b6a3.png)
	
4. Selecione o arquivo .tlb e esta tudo pronto para começar a utilizar a dll

![4](https://user-images.githubusercontent.com/54732019/76355355-e9107900-62f2-11ea-81b0-7c9b3b590acb.png)


### Realizando assintatura de XML:

Para realizar uma assintatura de XML, você poderá utilizar a função **assinarXML** da dll. Veja abaixo sobre os parâmetros necessários, e um exemplo de chamada do método.

#### Parâmetros:

Parametros    | Tipo de Dado | Descrição
:------------:|:------------:|:-----------
conteudo      | String       | O documente XML a ser assinado.
refURI        | String       | Nodo do XML que deve assinar. Ex.: <ul> <li>**infNFe** - NFe</li> <li>**infNFe** - NFCe</li> <li>**infCTe** - CTe</li> <li>**infMDFe** - MDFe</li> <li>**infBPe** - BPe</li> </ul> .
cnpjEmitente  | String   	 | CNPJ do Emitente da nota para que seja encontrado o certificado.


#### Exemplo de chamada:

Após ter todos os parâmetros listados acima, você deverá fazer a chamada da função. Veja o código em VB6 de exemplo abaixo:
    
    Dim assinaturaXML2 As New assinaturaXML2.Principal
	Dim xmlAssinado As String

	xmlAssinado = assinaturaXML2.assinarXML(conteudo, "infNFe", "11111111111111")
    MsgBox (xmlAssinado)	

A função **assinarXML** fará a assinatura do XML e retornará o mesmo assinado em modo de String.