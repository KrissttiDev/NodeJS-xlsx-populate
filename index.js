import XlsxPopulate from "xlsx-populate";

async function main(){
    //const workbook = await XlsxPopulate.fromFileAsync("./salida3.xlsx");
    //const value = workbook.sheet("Sheet1").range('A1:B2').value();
    //const value = workbook.sheet("Sheet1").usedRange().value();
    //console.log(value);
    /*const value1 = workbook.sheet("Sheet1").cell('A1').value();
    const value2 = workbook.sheet("Sheet1").cell('A2').value();
    console.log(value1 + " " + value2);*/
    /*const workbook = await XlsxPopulate.fromBlankAsync();
    workbook.sheet(0).cell("A1").value("Nombre");
    workbook.sheet(0).cell("B1").value("Apellido");
    workbook.sheet(0).cell("C1").value("Edad");

    workbook.sheet(0).cell("A2").value("Juan");
    workbook.sheet(0).cell("B2").value("Perez");
    workbook.sheet(0).cell("C2").value(25);

    workbook.sheet(0).cell("A3").value("Maria");
    workbook.sheet(0).cell("B3").value("Gomez");
    workbook.sheet(0).cell("C3").value(30);


    workbook.toFileAsync("./salida3.xlsx");*/

    /*const workbook = await XlsxPopulate.fromBlankAsync();
    workbook.sheet(0).cell("A1").value([
        [new Date().getDate(),new Date().getMonth()+1,new Date().getFullYear()],
        ['Nombre','Apellido', 'Edad'],
        ['Juan', 'Perez', 25],
        ['Pedro', 'Gomez', 30],
        ['Maria', 'Gonzalex', 30]
    ]);
    workbook.toFileAsync("./salida4.xlsx")*/

    const workbook = await XlsxPopulate.fromFileAsync("./salida6.xlsx");
    //const sheet =workbook.sheet(0);
    //console.log(sheet.name());
    //workbook.addSheet("Hoja2");
    //workbook.toFileAsync("./salida5.xlsx");
    //console.log(workbook.sheets().map((sheet) => sheet.name()));
    //workbook.sheet('Hoja2').name('Hoja de Prueba');
    //workbook.toFileAsync("./salida6.xlsx");
    workbook.deleteSheet('Hoja de Prueba');
    //workbook.toFileAsync("./salida7.xlsx");
    workbook.toFileAsync("./swguro.xlsx",{
        password:"123456"
    });
}

main();