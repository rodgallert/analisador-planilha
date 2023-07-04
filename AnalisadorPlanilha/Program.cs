// See https://aka.ms/new-console-template for more information
using AnalisadorPlanilha.Services;

Console.WriteLine("Informe o diretorio contendo a lista de destinatarios");
string diretorioDestinatarios = Console.ReadLine();
Console.WriteLine("Informe o nome das colunas, separadas por virgula, com o email dos destinatarios");
string colunasDestinatarios = Console.ReadLine();
Console.WriteLine("Informe o nome da planilha contendo o e-mail dos destinatarios");
string planilhaDestinatarios = Console.ReadLine();

Console.WriteLine("Informe o diretorio contendo a lista de eventos a serem notificados");
string diretorioEventos = Console.ReadLine();
Console.WriteLine("Informe as colunas, separadas por virgula, a serem notificadas");
string colunasEventos = Console.ReadLine();
Console.WriteLine("Informe o nome da planilha contendo os eventos a serem notificados");
string planilhaEventos = Console.ReadLine();

Console.WriteLine("Aguarde processamento");

try
{
    ExcelService.ProcessarLista(diretorioDestinatarios, colunasDestinatarios, planilhaDestinatarios, diretorioEventos, colunasEventos, planilhaEventos);
}
catch (Exception ex)
{
    Console.WriteLine("Um erro ocorreu, tente novamente ou comunique o desenveolvedor:");
    Console.WriteLine(ex.Message);
}