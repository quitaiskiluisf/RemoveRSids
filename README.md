# RemoveRSids
Remove os RevisionSessionIdentifier (tags RSID e similares) de documentos do Microsoft Word no formato ".docx".

## Origem

O Microsoft Office, no seu formato [`.docx`](https://en.wikipedia.org/wiki/Office_Open_XML), armazena em
seus arquivos várias tags `<w:r w:rsidR="00FF1F2C">` junto ao conteúdo do texto, além de atributos `w:rsidRDefault`
incluídos em diversas tags. Essas tags [armazenam números aleatórios]
(https://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.rsid(v=office.14).aspx)
que servem para [melhorar a precisão do algoritmo de comparação/combinação de documentos]
(http://blogs.msdn.com/b/brian_jones/archive/2006/12/11/what-s-up-with-all-those-rsids.aspx).

Por requisitos da instituição onde fiz minha graduação em Engenharia de Produção (em 2012), o TFC (Trabalho Final de Curso)
deveria estar em formato ".doc" ou ".docx" e, por isso, fiz o trabalho utilizando o Microsoft Word 2010. Para facilitar o
gerenciamento, o trabalho estava em um repositório SVN, em que eu e o orientador tínhamos acesso.

Como o formato `.docx` é, na verdade, composto por vários arquivos `.xml` compactados em um único container, eu queria poder
comparar duas revisões de um arquivo simplesmente extraíndo o seu conteúdo em diretórios distintos e utilizar uma ferramenta
de comparação de arquivos (como o [WinMerge] (http://winmerge.org/), por exemplo) para obter as diferenças. Entretanto,
com todos esses `rsidR`, na maior parte das vezes, o resultado dessa comparação ficava muito confuso para poder ser utilizado.

Assim, esse aplicativo surgiu com esse propósito: abrir um documento do Microsoft Word em formato `.docx` e eliminar todos os
`rsid` encontrados nele.

## Linguagem

Esse aplicativo foi criado utilizando-se o Microsoft Visual Studio 2010, versão Express.

## Uso

Após compilado, o aplicativo é chamado através do terminal da seguinte forma:

```
RemoveRSids.exe arquivo-exemplo.docx
```

Ao término da execução, serão gerados os seguintes arquivos:
* arquivo-exemplo.docx - arquivo com os `rsid` removidos.
* arquivo-exemplo.docx.bak - arquivo original, renomeado.

## Licença

Este aplicativo foi disponibilizado no domínio público sob os termos da licença [CC0]
(https://creativecommons.org/publicdomain/zero/1.0/), podendo, assim, ser utilizado, modificado, redistribuído e suas
partes utilizadas para qualquer fim, inclusive fins comerciais.
