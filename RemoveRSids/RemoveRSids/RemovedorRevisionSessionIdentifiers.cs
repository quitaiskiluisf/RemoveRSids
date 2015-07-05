using System;
using System.Collections;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Xml;

namespace RemoveRSids
{
    /// <summary>
    /// Classe responsável por fazer o processamento de um arquivo .docx, removendo todos os Revision Session Identifiers encontrados.
    /// </summary>
    public class RemovedorRevisionSessionIdentifiers: IDisposable
    {
        /// <summary>
        /// StreamWriter para onde serão enviados as mensagens de log.
        /// A stream onde o log será gravado é informada ao criar a classe (como um parâmetro opcional)
        /// Caso ele não seja informado, nenhum log será gerado
        /// </summary>
        public StreamWriter DestinoLog { get; private set; }


        /// <summary>
        /// Permite que seja recuperado o nome do arquivo que a classe irá processar
        /// </summary>
        public string Arquivo { get; private set; }


        /// <summary>
        /// Define o arquivo temporário, onde as alterações estão sendo feitas
        /// </summary>
        private string ArquivoTemporario { get; set; }


        /// <summary>
        /// 
        /// </summary>
        public readonly string[] AtributosAExcluir = new string[] { "w:rsidR", "w:rsidRDefault", "w:rsidP", "w:rsidSect", "w:rsidRPr", "w:rsidTr" };


        /// <summary>
        /// 
        /// </summary>
        public readonly string[] AtributosAIgnorar = new string[] { };


        /// <summary>
        /// 
        /// </summary>
        public readonly string[] AtributosSuspeitos = new string[] { "w:rsid" };


        /// <summary>
        /// 
        /// </summary>
        public readonly string[] TagsAExcluir = new string[] { "w:rsid", "w:rsids" };


        /// <summary>
        /// 
        /// </summary>
        public readonly string[] TagsAIgnorar = new string[] { };


        /// <summary>
        /// 
        /// </summary>
        public readonly string[] TagsSuspeitas = new string[] { "w:rsid" };


        /// <summary>
        /// Registra uma mensagem no log (caso foi informado um destino para o log ao instanciar a classe)
        /// </summary>
        /// <param name="mensagem">Mensagem a ser registrada</param>
        private void RegistraLog(string mensagem)
        {
            //Verifica se o log deve ser gerado
            if (this.DestinoLog != null)
            {
                this.DestinoLog.WriteLine(mensagem);
                this.DestinoLog.Flush();
            }
        }


        /// <summary>
        /// Registra uma mensagem no log (caso foi informado um destino para o log ao instanciar a classe.
        /// </summary>
        /// <param name="mensagem">Mensagem a ser registrada</param>
        /// <param name="substituicoes">Parâmetros que serão substituídos na mensaem de log (utilizando-se System.Format())</param>
        private void RegistraLog(string mensagem, params string[] substituicoes)
        {
            this.RegistraLog(String.Format(mensagem, substituicoes));
        }


        /// <summary>
        /// Inicia o processamento do arquivo.
        /// Ao final, o arquivo original possuirá o sufixo ".bak" e o arquivo alterado irá conter o nome original
        /// </summary>
        public void Processa()
        {
            //Cria uma cópia do arquivo em um diretório temporário
            this.ArquivoTemporario = Path.GetTempFileName();
            this.RegistraLog("Criando uma cópia do arquivo sob o nome temporário {0}", this.ArquivoTemporario);
            using (Stream novo = new FileStream(this.ArquivoTemporario, FileMode.Open, FileAccess.ReadWrite))
            {
                using (Stream antigo = new FileStream(this.Arquivo, FileMode.Open, FileAccess.Read))
                    antigo.CopyTo(novo);

                this.RegistraLog("Abrindo o arquivo temporário");
                novo.Seek(0, SeekOrigin.Begin);
                Package docx = ZipPackage.Open(novo, FileMode.Open, FileAccess.ReadWrite);
                
                this.RegistraLog("Percorrendo o conteúdo do arquivo");
                foreach (PackagePart arquivo in docx.GetParts())
                    this.ProcessaPackagePart(arquivo);

                this.RegistraLog("Fechando o arquivo");
                docx.Close();
            }
            Rotate();
        }


        /// <summary>
        /// Recebe um arquivo e decide o que fazer com ele
        /// </summary>
        /// <param name="arquivo">Arquivo sob o qual será tomada a decisão</param>
        private void ProcessaPackagePart(PackagePart arquivo)
        {
            string NomeArquivo = arquivo.Uri.ToString();
            this.RegistraLog("Extraindo o conteúdo do arquivo \"{0}\"", NomeArquivo);
            Stream ConteudoArquivo = arquivo.GetStream();
            this.RegistraLog("Tamanho do arquivo extraido: {0} bytes", ConteudoArquivo.Length.ToString());
            if (!VerificaProcessarArquivo(arquivo.Uri.ToString()))
                this.RegistraLog("Não irá processar o arquivo \"{0}\"", NomeArquivo);
            else
            {
                this.RegistraLog(String.Format("Processando o arquivo \"{0}\"", NomeArquivo));
                XmlDocument xml = new XmlDocument();
                ConteudoArquivo.Seek(0, SeekOrigin.Begin);
                xml.Load(ConteudoArquivo);
                //Faz o processamento e, caso foram feitas modificações, substitui o conteúdo original
                //do arquivo pelo arquivo modificado
                bool ResultadoRSids = ProcessaXmlRsids(xml);
                bool ResultadoUnificacao = ProcessaXmlUnificaRevisoes(xml);
                if (ResultadoRSids || ResultadoUnificacao)
                {
                    //Deixa o XML igual ao gerado pelo word (com a primeira linha contendo o <?xml ... ?> e a segunda contendo o conteudo do xml
                    //e também elimina o espaço existente entre o último argumento e a "/>". Já coloca o resultado em um array de bytes, prontos
                    //para serem gravados na stream. Essa parte tem que ser melhorada, porque essa não é a melhor forma de fazer isso (é meio porca)
                    byte[] xmlutf8 = Encoding.UTF8.GetBytes(xml.InnerXml.Replace("?>", "?>\n").Replace(" />", "/>"));
                    //Limpa o conteúdo da stream (que, no momento, contém o arquivo original), e escreve nela o conteúdo do arquivo novo
                    this.RegistraLog("Houveram alterações no arquivo \"{0}\"... fazendo a substituição pela versão nova", NomeArquivo);
                    ConteudoArquivo.SetLength(0);
                    ConteudoArquivo.Write(xmlutf8, 0, xmlutf8.Length);
                }
            }
        }

        /// <summary>
        /// Traz o arquivo temporário para o diretório atual, criando um backup do arquivo original
        /// </summary>
        private void Rotate()
        {
            FileInfo ArquivoTemporario = new FileInfo(this.ArquivoTemporario);
            string ArquivoBackup = this.Arquivo;
            while (File.Exists(ArquivoBackup))
                ArquivoBackup += ".bak";
            this.RegistraLog("Criando backup do arquivo original (sob o nome \"{0}\") e substituindo-o pelo arquivo novo", this.Arquivo);
            ArquivoTemporario.Replace(this.Arquivo, ArquivoBackup);
        }

        /// <summary>
        /// Recebe um XMLDocument e o varre em busca dos famigerados rsids. Ao encontrar um elemento desses, o remove
        /// </summary>
        /// <param name="xml">Xml onde o parâmetro será procurado</param>
        /// <returns>True caso houveram alterações; false caso não houveram</returns>
        protected bool ProcessaXmlRsids(XmlDocument xml)
        {
            return ProcessaXmlRsidsRecursivo((XmlNode)xml.DocumentElement);
        }


        /// <summary>
        /// Recebe um XMLDocument e o varre em busca dos famigerados rsids. Ao encontrar um elemento desses, o remove
        /// </summary>
        /// <param name="xml">Xml onde o parâmetro será procurado</param>
        /// <returns>True caso houveram alterações neste node ou em algum dos filhos; false caso não houveram</returns>
        protected bool ProcessaXmlRsidsRecursivo(XmlNode xml)
        {
            bool resultado = false;

            //Tratamento de tags a serem excluídas
            if (VerificaExcluiTag(xml))
            {
                //Se o xmlnode deve ser excluído, o exclui e já retorna o controle (já que, como ele foi excluído, não há mais nada a ser feito
                //com ele ou com seus descendentes
                xml.ParentNode.RemoveChild(xml);
                return true;
            }
            else if (VerificaTagSuspeita(xml) && !VerificaIgnoraTag(xml))
                throw new Exception(String.Format("A tag \"{0}\" é suspeita e não se encontra na lista de exclusões. Verificar o que deve ser feito com ela", xml.Name));

            //Tratamento de atributos a serem excluídos
            //Verifica se ele possui atributos (TextNodes, por exemplo, não possuem)
            if (xml.Attributes != null)
            {
                //Primeiro, copia em um array os nomes dos atributos (não é possível pegar direto porque ele dá erro no Enumerator quando
                //se remove um item (o que faz bastante sentido...)
                ArrayList ListaAtributos = new ArrayList();
                foreach (XmlAttribute attr in xml.Attributes)
                    ListaAtributos.Add(attr.Name);

                foreach (string attr in ListaAtributos)
                    if (VerificaExcluiAtributo(attr))
                    {
                        //Remove o atributo e sinaliza que houveram alterações
                        xml.Attributes.RemoveNamedItem(attr);
                        resultado = true;
                    }
                    else if (VerificaAtributoSuspeito(attr) && !VerificaIgnoraAtributo(attr))
                        throw new Exception(String.Format("O atributo \"{0}\" (na tag \"{1}\") é suspeito e não se encontra na lista de exclusões. Verificar o que deve ser feito com ele.", attr, xml.Name));
            }

            //Processa os filhos
            foreach (XmlNode filho in xml.ChildNodes)
                if (ProcessaXmlRsidsRecursivo(filho)) resultado = true;

            return resultado;
        }


        /// <summary>
        /// Faz as verificações necessárias para garantir que as revisões informadas poderão ser combinadas.
        /// Elas somente serão combinadas se estiverem em sequencia (ou seja, se não houver nenhum outro elemento entre elas)
        /// </summary>
        /// <param name="r1">Revisão 1</param>
        /// <param name="r2">Revisão 2</param>
        /// <returns>True caso poderão ser combinadas; senão false</returns>
        private static bool VerificaPodeUnificarRevisoes(XmlNode r1, XmlNode r2)
        {
            return (r1.NextSibling == r2 || r2.NextSibling == r1)
                   && r1.HasChildNodes
                   && r1.FirstChild.Name == "w:t"
                   && r1.FirstChild.HasChildNodes
                   && r1.FirstChild.FirstChild is XmlText
                   && r2.HasChildNodes
                   && r2.FirstChild.Name == "w:t"
                   && r2.FirstChild.HasChildNodes
                   && r2.FirstChild.FirstChild is XmlText;
        }


        /// <summary>
        /// Dadas 2 revisões, realiza o processo de unificação entre elas, caso elas sejam siblings
        /// </summary>
        /// <param name="r1">Primeira revisão</param>
        /// <param name="r2">Segunda revisão</param>
        private static void UnificaRevisoes(XmlNode r1, XmlNode r2)
        {
            //Determina qual node aparece antes, se o r1 ou o r2. Isso porque a tag que vem depois será removida da árvore,
            //enquanto que o seu conteúdo será acrescentado à tag que vem primeiro.
            XmlNode antes, depois;

            //Garante que eles são siblings
            if (r1.NextSibling == r2 || r2.NextSibling == r1)
            {
                bool EmOrdem = r1.NextSibling == r2;
                antes = (EmOrdem ? r1 : r2);
                depois = (EmOrdem ? r2 : r1);
            }
            else
                throw new Exception("Não é possível unificar as tags informadas porque elas não são sibblings");

            //Junta os textos
            XmlNode WTAntes = antes.FirstChild;
            XmlText TextoAntes = (XmlText)WTAntes.FirstChild;
            XmlText TextoDepois = (XmlText)depois.FirstChild.FirstChild;

            TextoAntes.AppendData(TextoDepois.Data);

            //Verifica se o atributo xml:space=preserve na tag w:t deve ser acrescentado, mantido ou eliminado
            const string ATTR_XML_SPACE = "xml:space";
            bool Existe = WTAntes.Attributes.GetNamedItem(ATTR_XML_SPACE) != null;
            bool DeveExistir = (TextoAntes.Data.StartsWith(" ") || TextoAntes.Data.EndsWith(" "));

            if (DeveExistir && !Existe)
                ((XmlElement)WTAntes).SetAttribute(ATTR_XML_SPACE, "preserve");
            else if (!DeveExistir && Existe)
                WTAntes.Attributes.RemoveNamedItem(ATTR_XML_SPACE);

            //Elimina o segundo node (pois ele foi unificado com o primeiro
            depois.ParentNode.RemoveChild(depois);
        }


        /// <summary>
        /// Procura por tags w:r que são sibblings e as remove
        /// </summary>
        /// <param name="xml">Tag onde serão procurados os w:r</param>
        /// <returns>True caso houveram alterações; false caso não houveram</returns>
        protected bool ProcessaXmlUnificaRevisoes(XmlDocument xml)
        {
            //return ProcessaXmlUnificaRevisoesRecursivo((XmlNode)xml.DocumentElement);
            XmlNodeList lista = xml.GetElementsByTagName("w:r");

            //Percorre a lista de trás para frente (para evitar que a alteração no número de nodes influencie na
            //lógica de cálculo
            for (int i = lista.Count - 1; i > 0; i--)
            {
                XmlNode rAtual = lista[i];
                XmlNode rAnterior = lista[i - 1];

                if (VerificaPodeUnificarRevisoes(rAtual, rAnterior))
                {
                    RegistraLog("Unificando uma revisão...");
                    UnificaRevisoes(rAtual, rAnterior);
                }
            }
            return true;
        }


        /// <summary>
        /// Verifica algum dos itens da "lista" inicia com o texto informado em "item" (case insensitive e culture insensitive)
        /// </summary>
        /// <param name="item">Texto que será procurado no início das strings do array</param>
        /// <param name="lista">Array onde o texto será procurado</param>
        /// <returns>True caso ele foi encontrado; false caso não foi</returns>
        private static bool VerificaSeIniciaCom(string item, params string[] lista)
        {
            foreach (string a in lista)
                if (item.StartsWith(a, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            return false;
        }


        /// <summary>
        /// Verifica algum dos itens da "lista" é igual ao texto informado em "item" (case insensitive e culture insensitive)
        /// </summary>
        /// <param name="item">Texto que será procurado nas strings do array</param>
        /// <param name="lista">Array onde o texto será procurado</param>
        /// <returns>True caso ele foi encontrado; false caso não foi</returns>
        private static bool VerificaSeExiste(string item, params string[] lista)
        {
            foreach (string a in lista)
                if (item.Equals(a, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            return false;
        }


        /// <summary>
        /// Verifica se a tag atual encontra-se na lista de tags suspeitas (que iniciem com o texto informado)
        /// </summary>
        /// <param name="xml">XmlNode atual, onde serão verificadas a existência de tags suspeitas</param>
        /// <returns>True caso a tag é suspeita; senão false</returns>
        protected bool VerificaTagSuspeita(XmlNode xml)
        {
            return VerificaSeIniciaCom(xml.Name, this.TagsSuspeitas);
        }


        /// <summary>
        /// Verifica se a tag atual deve ser excluída, de acordo com as regras
        /// </summary>
        /// <param name="xml">XmlNode atual, que deverá ser verificado</param>
        /// <returns>True caso ele foi excluído; senão false</returns>
        protected bool VerificaExcluiTag(XmlNode xml)
        {
            return VerificaSeExiste(xml.Name, this.TagsAExcluir);
        }


        /// <summary>
        /// Verifica se a tag atual encontra-se na lista de tags a ignorar, de acordo com as regras
        /// </summary>
        /// <param name="xml">XmlNode atual, que deverá ser verificado</param>
        /// <returns>True caso ela deve ser ignorada; senão false</returns>
        protected bool VerificaIgnoraTag(XmlNode xml)
        {
            return VerificaSeExiste(xml.Name, this.TagsAIgnorar);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="attr">Atributo atual, que será verificado contra a lista de tags suspeitas</param>
        /// <returns>True caso o atributo é suspeito; senão false</returns>
        protected bool VerificaAtributoSuspeito(string attr)
        {
            return VerificaSeIniciaCom(attr, this.AtributosSuspeitos);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="attr">Atributo atual, que será verificado contra a lista de atributos a remover</param>
        /// <returns></returns>
        protected bool VerificaExcluiAtributo(string attr)
        {
            return VerificaSeExiste(attr, this.AtributosAExcluir);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="attr">Atributo atual, que será verificado contra a lista de atributos a ignorar</param>
        /// <returns></returns>
        protected bool VerificaIgnoraAtributo(string attr)
        {
            return VerificaSeExiste(attr, this.AtributosAIgnorar);
        }


        /// <summary>
        /// Verifica se o arquivo informado deverá ser processado ou não
        /// </summary>
        /// <param name="Nome">Nome do arquivo</param>
        /// <returns>Resultado da operação</returns>
        private static bool VerificaProcessarArquivo(string Nome)
        {
            if (Nome.EndsWith(".xml", StringComparison.InvariantCultureIgnoreCase))
                return true;
            return false;
        }


        /// <summary>
        /// Faz as validações necessárias para garantir que foi informado um arquivo válido, retornando
        /// o resultado da validação
        /// </summary>
        /// <param name="Documento">Arquivo que deverá ser validado</param>
        /// <returns>O resultado da validação</returns>
        public static bool ArquivoValido(string Documento)
        {
            return true;
        }


        /// <summary>
        /// 
        /// </summary>
        private void IniciaLog(Stream destino)
        {
            if (destino != null)
                this.DestinoLog = new StreamWriter(destino, Encoding.UTF8);
        }


        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            this.DestinoLog.Flush();
            this.DestinoLog.Dispose();
        }


        /// <summary>
        /// Constructor da classe
        /// </summary>
        /// <param name="Documento">Nome do documento original (de onde serão extraídos os rsids)</param>
        /// <param name="Log">Stream para onde serão enviadas as mensagens de log</param>
        public RemovedorRevisionSessionIdentifiers(string Documento, Stream Destino)
            : base()
        {
            this.IniciaLog(Destino);
            this.RegistraLog(String.Format("Iniciando o carregamento do arquivo \"{0}\"", Documento));
            if (!File.Exists(Documento))
                throw new FileNotFoundException(String.Empty, Documento);
            if (!ArquivoValido(Documento))
                throw new Exception("O documento informado não é um documento válido.");
            this.Arquivo = Documento;
        }
    }
}
