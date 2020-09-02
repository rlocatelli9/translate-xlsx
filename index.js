const excelToJson = require('convert-excel-to-json');
var fs = require('fs');
const path = require("path");
const {format} = require('date-fns');

const oneDay = 24 * 60 * 60 * 1000; // cálculo em timestamp de 1 dia
const thirtyDay = oneDay * 30;  // cálculo em timestamp de 30 dias
let result = {};
let arrayConfig = [];

//lendo arquivo .json no repositorio raiz
console.log("[DEBUG] Lendo arquivo de configuração...")
let settingsFile = fs.readFileSync('settings.json', (err, data) => {
  if (err) throw new Error('Erro ao ler arquivo de configuracao...', err);
  return data;
});
settingsFile = JSON.parse(settingsFile);

const destination = path.join(__dirname, settingsFile.destinationFile);

try {
  console.log("[DEBUG] Interpretando arquivo xlsx...")
  /**
   * Função responsável pela interpretação do xlsx
   * @params config: {}, sourceFile: any
   * @returns {}
   */
  result = excelToJson({
    sourceFile: settingsFile.sourceFile,
    header:{
      rows: 1
    },
    columnToKey: settingsFile.columnToKey,
    sheets: settingsFile.sheets
  });
} catch (error) {
  console.error('[ERROR] Falha ao interpretar arquivo xlsx.', error);
  throw new Error('Falha na interpretação do arquivo');
}

const arrayData = result.Dados; // passando array dos dados interpretados

/**
 * Fazendo tratativa dos dados interpretados do arquivo origem
 * Tratativa feita de acordo com o padrão necessário para o arquivo de saída
 */
arrayData.map(element => {
  if(element.Contrato !== ''){
    const contrato = element.Contrato.split('.').join('').split('/').join('');
    // console.debug('CONTRATO: ', contrato);
    element.Contrato = ('00000000000000000000' + (contrato)).slice(-20);
    // element.Contrato = contrato;
  }else{
    element.Contrato = '00000000000000000000';
  }
  element.dateInitial = format(new Date().getTime() - (thirtyDay * element.NumPrestacao), 'ddMMyyyy');
  element.dateFinal = format(new Date().getTime() + (thirtyDay * element.PrazoRemanescente), 'ddMMyyyy');
  element.Matricula = '                    ';
  if(element.Nome !== '' && element.Nome !== '.'){
    let lengthName = element.Nome.length;
    let contador = 0;
    let nameSplited = "";

    if(lengthName > 30){
      while (contador < 30) {
        nameSplited += element.Nome[contador];
        contador ++;
      }
      element.Nome = nameSplited;
      console.log("NOME CORTADO: ", element.Nome);
    } else {
      while (lengthName < 30) {
        element.Nome += ' ';
        lengthName = element.Nome.length;
      }
      console.log("NOME: ", element.Nome);
    }
  } else {
    console.error('[ERROR] Nome fora do padrão.');
    throw new Error('Arquivo fora do padrão definido! Verifique o campo NOME.')
  }
  if(element.Cpf !== ''){
    const cpf = element.Cpf.split('.').join('').split('-').join('');
    if(cpf.length === 11) element.Cpf = cpf;
    else {
      console.error('[ERROR] CPF fora do padrão.');
      throw new Error('Arquivo fora do padrão definido! Verifique o campo CPF.')
    }
  } else {
    console.error('[ERROR] CPF vazio.');
    throw new Error('Arquivo fora do padrão definido! Verifique o campo CPF.')
  }
  if(element.ValorPrestacao !== ''){
    console.log('Valor prestacao original: ', element.ValorPrestacao.toFixed(2));
    const valorPrestacao = element.ValorPrestacao.toFixed(2).split(',').join('').split('.').join('');
    // console.debug('VALOR PRESTACAO: ', valorPrestacao);
    console.debug('VALOR prestacao formatado: ', valorPrestacao);
    element.ValorPrestacao = ('000000000000000' + (valorPrestacao)).slice(-15);
  }else{
    element.ValorPrestacao = '000000000000000';
  }
  if(element.ValorPagar !== ''){
    // console.log('Valor original: ', element.ValorPagar.toFixed(2));
    const valorPagar = element.ValorPagar.toFixed(2).split(',').join('').split('.').join('');
    // console.debug('VALOR formatado: ', valorPagar);
    element.ValorPagar = valorPagar;
  }else{
    element.ValorPagar = '000000000000000';
  }
  if(element.SituacaoDesconto !== ''){
    const situacaoDesconto = element.SituacaoDesconto.split('-');
    // console.debug('CODIGO SITUACAO: ', situacaoDesconto);
    element.SituacaoDesconto = situacaoDesconto[0];
    element.SituacaoDesconto = element.SituacaoDesconto.toString();

    let lengthSituacaoDesconto = element.SituacaoDesconto.length;
    let contador = 0;
    let situacaoDescontoSplited = "";

    if(lengthSituacaoDesconto > 20){
      while (contador < 20) {
        situacaoDescontoSplited += element.SituacaoDesconto[contador];
        contador ++;
      }
      element.SituacaoDesconto = situacaoDescontoSplited;
      // console.log("NOME CORTADO: ", element.Nome);
    } else {
      while (lengthSituacaoDesconto < 20) {
        element.SituacaoDesconto += ' ';
        lengthSituacaoDesconto = element.SituacaoDesconto.length;
      }
      // console.log("NOME: ", element.Nome);
    }
  }else{
    element.SituacaoDesconto = '00000000000000000000';
  }
  if(element.PrazoTotal !== ''){
    element.PrazoTotal = element.PrazoTotal.toString();
    const prazoLength = element.PrazoTotal.length;
    if(prazoLength < 3) {
      element.PrazoTotal = ('000' + (element.PrazoTotal)).slice(-3);
    }
  }else{
    element.PrazoTotal = "000";
  }
  element.ValorContratado = ('000000000000000' + ((Number(element.ValorPrestacao) * Number(element.PrazoTotal)))).slice(-15);
  element.margemIncidente = "EMPRESTIMO          ";
  arrayConfig = [...arrayConfig, `${element.Cpf}${element.Nome}${element.Matricula}              ${element.SituacaoDesconto}EMPRÉSTIMO          ${element.dateInitial}${element.dateFinal}N${element.PrazoTotal}${element.ValorPrestacao}${element.ValorContratado}${element.Contrato}${element.margemIncidente}\n`];
  // arrayConfig = [...arrayConfig, `${element.Cpf}${element.Nome}${element.Matricula}${element.SituacaoDesconto}EMPRÉSTIMO    ${element.dateInitial}${element.dateFinal}N${element.PrazoTotal}${element.ValorPrestacao}${element.ValorContratado}${element.Contrato}${element.margemIncidente}\n`];
})

try {
  /**
   * Escrevendo no arquivo cada resgistro tratado, linha por linha
   */
  arrayConfig.forEach(line => {
    fs.appendFileSync(
      destination,
      line.toString(),
      "utf-8"
    );
  });
  console.debug('[DEBUG] Arquivo TXT gerado!');
} catch (error) {
  console.error('[ERROR] Falha ao criar arquivo txt.', error);
  throw new Error('Falha na criação do arquivo txt');
}
