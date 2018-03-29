using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using Interop_Word = Microsoft.Office.Interop.Word;

namespace GeradorBookServico
{
    public partial class Form1 : Form
    {
        public string CaminhoDiretorioCsfServicoPagamento = @"C:\CSF\17Dez01\Financeiro\Servicos\Web\Pagamentos\Csf.Fin.Svc.Pagamentos\bin";
                
        public Form1()
        {
            InitializeComponent();
        }

        public List<string> ObterArquivos(string diretorio)
        {
            List<string> arquivos;
            List<string> arquivosSomenteNome;

            arquivos = Directory.GetFiles(diretorio, "*.dll", SearchOption.AllDirectories).ToList();

            arquivosSomenteNome = new List<string>();

            foreach (string  arquivo in arquivos)
            {
                arquivosSomenteNome.Add(Path.GetFileName(arquivo));
            }

            return arquivosSomenteNome;
        }

        public void Gerar(string endereco, ref Interop_Word.Document documento) 
        {
            Assembly amostraAssembly;

            //Carrega o assembly
            amostraAssembly = Assembly.LoadFrom(endereco);

            //Monta o nome do serviço
            documento.Content.SetRange(0, 0);
            documento.Content.Text = string.Format("Catalago de Serviços - {0}{1}", amostraAssembly.GetTypes()[0].FullName, Environment.NewLine);
            
            this.ObterMetodos(amostraAssembly.GetTypes()[0].GetMethods().ToList(), ref documento);
        }

        public void ObterMetodos(List<MethodInfo> metodos, ref Interop_Word.Document documento)
        {
            foreach (MethodInfo metodo in metodos)
            {
                Interop_Word.Paragraph descricaoMetodo = documento.Content.Paragraphs.Add();

                descricaoMetodo.Range.Text = metodo.Name;
                descricaoMetodo.Range.InsertParagraphAfter();

                this.ObterParametros(metodo.GetParameters().ToList(), ref documento, ref descricaoMetodo);
            }
        }

        public void ObterParametros(List<ParameterInfo> parametros, ref Interop_Word.Document documento, ref Interop_Word.Paragraph descricaoMetodo)
        {
            int indice = 1;

            Interop_Word.Table tabelaParametros = documento.Tables.Add(descricaoMetodo.Range, parametros.Count() + 1, 5);
            tabelaParametros.Borders.Enable = 1;

            EscreverHeader(ref documento, ref tabelaParametros);

            foreach (ParameterInfo parametro in parametros)
            {
                indice++;
                this.EscreverCorpo(parametro, indice, ref documento, ref tabelaParametros);

                //this.ObterPropriedades(((System.Reflection.TypeInfo)parametro.ParameterType).GetProperties().ToList(), ref documento, ref descricaoMetodo);
            }            
        }

        private void EscreverHeader(ref Interop_Word.Document documento, ref Interop_Word.Table tabelaParametros)
        {
            foreach (Interop_Word.Row row in tabelaParametros.Rows)
            {
                foreach (Interop_Word.Cell cell in row.Cells)
                {
                    if (cell.RowIndex == 1)
                    {
                        #region [ HEADER ]

                        switch (cell.ColumnIndex)
                        {
                            case 1:
                                cell.Range.Text = "Elemento";
                                break;
                            case 2:
                                cell.Range.Text = "Tipo";
                                break;
                            case 3:
                                cell.Range.Text = "Tamanho";
                                break;
                            case 4:
                                cell.Range.Text = "Mandatório";
                                break;
                            case 5:
                                cell.Range.Text = "Valor / Descrição";
                                break;
                            default:
                                cell.Range.Text = "--";
                                break;
                        }

                        cell.Range.Font.Bold = 1;

                        cell.Range.Font.Name = "verdana";
                        cell.Range.Font.Size = 10;

                        cell.Shading.BackgroundPatternColor = Interop_Word.WdColor.wdColorGray25;

                        cell.VerticalAlignment = Interop_Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = Interop_Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        #endregion
                    }
                }
            }
        }

        private void EscreverCorpo(ParameterInfo parametro, int indice, ref Interop_Word.Document documento, ref Interop_Word.Table tabelaParametros)
        {
            foreach (Interop_Word.Row row in tabelaParametros.Rows)
            {
                foreach (Interop_Word.Cell cell in row.Cells)
                {
                    #region [ REGISTROS ]

                    if (cell.RowIndex == indice)
                    {
                        switch (cell.ColumnIndex)
                        {
                            case 1:
                                cell.Range.Text = parametro.Name;
                                break;
                            case 2:
                                cell.Range.Text = parametro.ParameterType.Name;
                                break;
                            case 3:
                                cell.Range.Text = "--";
                                break;
                            case 4:
                                cell.Range.Text = (parametro.IsOptional ? "Não" : "Sim");
                                break;
                            case 5:
                                cell.Range.Text = parametro.HasDefaultValue ? parametro.DefaultValue.ToString() : "";
                                break;
                            default:
                                cell.Range.Text = "--";
                                break;
                        }
                    }

                    #endregion                    
                }
            }
        }

        public void ObterPropriedades(List<PropertyInfo> propriedades, ref Interop_Word.Document documento, 
            ref Interop_Word.Paragraph descricaoMetodo)
        {
            Interop_Word.Table tabelaParametros = documento.Tables.Add(descricaoMetodo.Range, propriedades.Count() + 1, 5);
            tabelaParametros.Borders.Enable = 1;
            PropertyInfo propriedade = propriedades.FirstOrDefault();
            int indice = 0;

            int linhas = tabelaParametros.Rows.Count;

            foreach (Interop_Word.Row row in tabelaParametros.Rows)
            {
                if (row.Index > 1)
                {
                    propriedade = indice < propriedades.Count() ? propriedades[indice] : null;
                    indice++;
                }

                foreach (Interop_Word.Cell cell in row.Cells)
                {
                    //if (propriedades != null)
                    if (propriedade != null)
                    {
                        if (cell.RowIndex == 1)
                        {
                            switch (cell.ColumnIndex)
                            {
                                case 1:
                                    cell.Range.Text = "Elemento";
                                    break;
                                case 2:
                                    cell.Range.Text = "Tipo";
                                    break;
                                case 3:
                                    cell.Range.Text = "Tamanho";
                                    break;
                                case 4:
                                    cell.Range.Text = "Mandatório";
                                    break;
                                case 5:
                                    cell.Range.Text = "Valor / Descrição";
                                    break;
                                default:
                                    cell.Range.Text = "--";
                                    break;
                            }

                            cell.Range.Font.Bold = 1;

                            cell.Range.Font.Name = "verdana";
                            cell.Range.Font.Size = 10;

                            cell.Shading.BackgroundPatternColor = Interop_Word.WdColor.wdColorGray25;

                            cell.VerticalAlignment = Interop_Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = Interop_Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                        else
                        {
                            switch (cell.ColumnIndex)
                            {
                                case 1:
                                    cell.Range.Text = propriedade.Name;
                                    break;
                                case 2:
                                    cell.Range.Text = propriedade.PropertyType.IsGenericType 
                                        ? propriedade.PropertyType.GetGenericArguments().Single().Name 
                                        : propriedade.PropertyType.Name;
                                    break;
                                //case 3:
                                //    cell.Range.Text = "VERIFICAR COMO PEGAR O TAMANHO DOS CAMPOS";
                                //    break;
                                case 4:
                                    cell.Range.Text = "pegar obrigatoriedade de propriedade";//(propriedade.IsOptional ? "C" : "M");
                                    break;
                                case 5:
                                    cell.Range.Text = "Valor / Descrição";
                                    break;
                                default:
                                    cell.Range.Text = "--";
                                    break;
                            }
                        }
                    }
                    else
                    {
                        //TODO: Caio - E se for nulo????
                    }
                }
            }
        }

        public void CarregarArquivosParaLeitura()
        {
            List<string> arquivos;

            arquivos = this.ObterArquivos(CaminhoDiretorioCsfServicoPagamento);

            checkedListBox_Arquivos.Items.AddRange(arquivos.ToArray());
        }

        public void EscolherPastaParaLeitura()
        {
            FolderBrowserDialog browser = new FolderBrowserDialog();

            DialogResult resultado = browser.ShowDialog();

            if (resultado.Equals(DialogResult.OK))
            {
                CaminhoDiretorioCsfServicoPagamento = browser.SelectedPath;
            }
        }

        public string EscolherPastaParaSalvar()
        {
            string pastaSalvar = string.Empty;
            FolderBrowserDialog browser = new FolderBrowserDialog();

            DialogResult resultado = browser.ShowDialog();

            if (resultado.Equals(DialogResult.OK))
            {
                pastaSalvar = browser.SelectedPath;
            }

            return pastaSalvar;
        }

        private void btnGerar_Click(object sender, EventArgs e)
        {
            foreach (string item in checkedListBox_Arquivos.CheckedItems)
            {
                try
                {
                    string pastaSalvar = string.Empty;
                    string nomeArquivo = Path.GetFileNameWithoutExtension(item) + DateTime.Now.ToString("HHmm_ddMMyyyy") + ".docx";
                    Interop_Word.Application documentoWord;
                    documentoWord = new Interop_Word.Application();

                    documentoWord.Visible = false;

                    Interop_Word.Document documento = documentoWord.Documents.Add();


                    documento.Content.SetRange(0, 0);

                    this.Gerar(string.Format(@"{0}\{1}", CaminhoDiretorioCsfServicoPagamento, item), ref documento);
                    
                    pastaSalvar = this.EscolherPastaParaSalvar();

                    if (string.IsNullOrEmpty(pastaSalvar))
                    {
                        pastaSalvar = @"C:\Temp";
                    }

                    documento.SaveAs(string.Format(@"{0}\{1}", pastaSalvar, nomeArquivo));
                    documento.Close();
                    documentoWord.Quit();
                    MessageBox.Show(string.Format("O arquivo [{0}] foi salvo no diretório [{1}].", nomeArquivo, pastaSalvar), "Salvo com sucesso!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocorreu um erro. " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.EscolherPastaParaLeitura();

            if (String.IsNullOrWhiteSpace(CaminhoDiretorioCsfServicoPagamento))
            {
                MessageBox.Show("Caminho de arquivos não selecionado!", "Erro no procedimento", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                this.CarregarArquivosParaLeitura();
            }
        }
    }
}
