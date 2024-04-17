const XLSX = require('xlsx');
const XLSX_CALC = require('xlsx-calc')
const fs = require('fs')
const path = require('path')
let formulajs = require('@formulajs/formulajs')

formulajs = { ...formulajs, ANIO: formulajs.YEAR }
console.log(typeof formulajs)
console.log('formulas', formulajs)

XLSX_CALC.import_functions(formulajs)

const year = formulajs.YEAR(45351)
console.log('YEAR', year)

const anio = formulajs.ANIO(45351)
console.log('ANIO', anio)

const AÑO = formulajs.YEAR
const anio2 = AÑO(45351)
console.log('ANIO', anio2)

const calculate = async (nameFile) => {
    const pathFile = path.join(path.dirname(__dirname), nameFile)
    const workbook = XLSX.readFile(pathFile);
    const sheetName = workbook.SheetNames[1]; // Hoja 2
    const worksheet = workbook.Sheets[sheetName];

    const cellRef = worksheet['B3'];
    const reference = XLSX.utils.format_cell(cellRef);
    console.log('!REF', reference);

    const cellSuma = worksheet['D6'];
    const suma = XLSX.utils.format_cell(cellSuma);
    console.log('SUMA', suma);

    const cellRedondeo = worksheet['C11'];
    const redondeo = XLSX.utils.format_cell(cellRedondeo);
    console.log('REDONDEO', redondeo); 

    const cellBuscarH = worksheet['B14'];
    const buscarH = XLSX.utils.format_cell(cellBuscarH);
    console.log('BUSCAR_H', buscarH); 

    const cellBuscarV = worksheet['B17'];
    const buscarV = XLSX.utils.format_cell(cellBuscarV);
    console.log('BUSCAR_V', buscarV);
    
    const cellEntero = worksheet['B20'];
    console.log('cellEntero', cellEntero)
    const entero = XLSX.utils.format_cell(cellEntero);
    console.log('ENTERO', entero);

    const cellAnio = worksheet['B23'];
    console.log('YEAR', cellAnio)
    const anio = XLSX.utils.format_cell(cellAnio);
    console.log('AÑO', anio);

    const demo1 = worksheet['B25'];
    const demo11 = XLSX.utils.format_cell(demo1);
    console.log('Demo1', demo11);

    const demo2 = worksheet['B26'];
    const demo22 = XLSX.utils.format_cell(demo2);
    console.log('Demo2', demo22);
}
// { t: 'n', v: 75, f: '+INT(L194/12)', w: '75' }
const writeData = async () => {
    const nameNewFile = 'Demo.xlsb'
    const workbook = XLSX.readFile('/Users/rogelioadriansuclupetello/Development/RIMAC/ApagadoFinrisk/examples/xlsx-poc/Libro.xlsx');
    let sheetName = workbook.SheetNames[0]; // Hoja 1
    let worksheet = workbook.Sheets[sheetName];
    worksheet['C3'].v = 800

    sheetName = workbook.SheetNames[1]; // Hoja 2
    worksheet = workbook.Sheets[sheetName];
    worksheet['A6'].v = 400
    worksheet['B6'].v = 500
    worksheet['A11'].v = 'NO'
    worksheet['A14'].v = 107
    worksheet['A17'].v = 100

    XLSX_CALC(workbook);

    XLSX.writeFileXLSX(workbook, nameNewFile);
    await calculate(nameNewFile)
    fs.unlinkSync(`/Users/rogelioadriansuclupetello/Development/RIMAC/ApagadoFinrisk/examples/xlsx-poc/${nameNewFile}`)
}

writeData()
