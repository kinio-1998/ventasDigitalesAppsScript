/*
  Proyecto: Ventas Digitales - Automatización con Google Apps Script
  Descripción: Este script automatiza la recopilación, análisis y visualización de ventas quincenales por mes a través de una hoja de cálculo.
  Autor: [Tu nombre]
  Fecha: [Fecha de creación]
*/

// Variables globales para manejo de datos
let datosFiltrados = [], rangoQuincenas = [], totalHojas = [];
let contador1ra = 0, total1ra = 0, contador2da = 0, total2da = 0;

// Referencia al archivo de Google Sheets activo y listado de hojas
const hojaActiva = SpreadsheetApp.getActiveSpreadsheet();
const hojas = new Object(hojaActiva.getSheets());

// Se filtran las hojas para excluir "ARTICULOS" y "Plantilla"
totalHojas.push(hojas.filter((hoja) => hoja.getName() != "ARTICULOS" && hoja.getName() != "Plantilla"));

// Lista de meses para análisis de fechas
const meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

// Función principal para actualizar los datos por hoja
const actualizar = () => {
  totalHojas[0].forEach((hoja) => {
    const datosVentas = hoja.getRange("A2:J").getValues();

    datosVentas.map((reg) => {
      if (reg[0] != '#N/A' && reg[0] != '' && reg[0] != ' ') {
        reg.push(hoja.getName()); // Se añade el nombre de la hoja al final
        datosFiltrados.push(reg);
      }
    });

    contar(hoja); // Llama a función para análisis por quincenas

    // Se imprimen los resultados en las celdas respectivas
    hoja.getRange("M2").activate().setValue(contador1ra).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID).setFontFamily('Roboto');
    hoja.getRange("M3").activate().setValue(contador2da).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID).setFontFamily('Roboto');
    hoja.getRange("N2").activate().setValue(total1ra).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID).setFontFamily('Roboto');
    hoja.getRange("N3").activate().setValue(total2da).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID).setFontFamily('Roboto');

    // Se reinician variables para siguiente iteración
    datosFiltrados.length = 0;
    contador1ra = 0;
    total1ra = 0;
    contador2da = 0;
    total2da = 0;
  });
};

// Función auxiliar para contar ventas en cada quincena
const contar = (hoja) => {
  const nombreMes = hoja.getName();
  const mesHoja = nombreMes.toUpperCase().split(' ');

  datosFiltrados.map((datos) => {
    if (datos[3] != "") {
      const fecha = datos[3];
      const mesIndex = meses.findIndex((element) => element == mesHoja[0]);

      if (fecha.getDate() >= 1 && fecha.getDate() <= 15 && fecha.getMonth() == mesIndex) {
        contador1ra += datos[4];
        total1ra += datos[9].toUpperCase() == "SHOPIFY" ? datos[7] : datos[5];
      } else if (fecha.getDate() > 15 && fecha.getDate() <= 31 && fecha.getMonth() == mesIndex) {
        contador2da += datos[4];
        total2da += datos[9].toUpperCase() == "SHOPIFY" ? datos[7] : datos[5];
      }
    }
  });
};

// Función para crear una nueva hoja a partir de la plantilla, el primer día del mes
const agregarHoja = () => {
  const plantilla = SpreadsheetApp.setActiveSheet(hojaActiva.getSheetByName('Plantilla'), true);
  const dia = new Date().getDate();
  const mes = new Date().getMonth();
  const anio = new Date().getFullYear();

  if (dia == 1) {
    plantilla.activate();
    hojaActiva.duplicateActiveSheet().showSheet();
    const nuevaHoja = SpreadsheetApp.setActiveSheet(hojaActiva.getSheetByName('Copia de Plantilla'), true);

    nuevaHoja.setName(meses[mes] + ' ' + anio);

    const valoresEstaticos = [["=LOOKUP(C2,ARTICULOS[SKU],ARTICULOS[NOMBRE])", "=LOOKUP(C2,ARTICULOS[SKU],ARTICULOS[PRECIO_TOTAL])", 0, '01/01/1900', 0, 0, 0, 0, "", ""], ["1era", 0, 0], ["2da", 0, 0], ["Total", "=SUM(M2:M3)", "=SUM(N2:N3)"]];

    nuevaHoja.getRange('A2:J2').activate().setValues([valoresEstaticos[0]]).setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment('center').setFontFamily('Roboto');
    nuevaHoja.getRange('A2').setHorizontalAlignment('left');

    nuevaHoja.getRange('L2:N2').activate().setValues([valoresEstaticos[1]]).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID).setFontFamily('Roboto');
    nuevaHoja.getRange('L3:N3').activate().setValues([valoresEstaticos[2]]).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID).setFontFamily('Roboto');
    nuevaHoja.getRange('L4:N4').activate().setValues([valoresEstaticos[3]]).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID).setFontFamily('Roboto');

    // Configuración visual de columnas
    hojaActiva.getActiveSheet().setColumnWidth(1, 350);
    hojaActiva.getActiveSheet().setColumnWidth(2, 120);
    hojaActiva.getActiveSheet().setColumnWidth(4, 115);
    hojaActiva.getActiveSheet().setColumnWidth(5, 135);
    hojaActiva.getActiveSheet().setColumnWidth(6, 110);
    hojaActiva.getActiveSheet().setColumnWidth(7, 135);
    hojaActiva.getActiveSheet().setColumnWidth(8, 130);
    hojaActiva.getActiveSheet().setColumnWidth(11, 150);
    hojaActiva.getActiveSheet().setColumnWidth(10, 130);
  }
};
