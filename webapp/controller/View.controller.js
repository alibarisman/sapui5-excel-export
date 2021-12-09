sap.ui.define([
	"sap/ui/core/mvc/Controller",
	"com/nttdata/ExcelJs/util/xlsxfullmin",
	"com/nttdata/ExcelJs/util/exceljs",
	"com/nttdata/ExcelJs/util/FileSaver"
], function (Controller, xlsxfullmin, exceljs, FileSaver) {
	"use strict";

	return Controller.extend("com.nttdata.ExcelJs.controller.View", {

		onInit: function () {

		},

		/* global XLSX:true */
		/* global saveAs:true */
		onExportSheetJs: function (data) {
			let oMockModel = this.getOwnerComponent().getModel("mock").getData();
			let mockArr = [];
			
			let wb = XLSX.utils.book_new();
			wb.Props = {
				Title: "Excel Document",
				Subject: "SheetJs",
				Author: "Ali Hulusi Barisman",
				CreatedDate: new Date()
			};

			for (let i = 0; i < oMockModel.requests.length; i++) {
				let row = {
					"Request": oMockModel.requests[i].inputPRAmount,
					"Material": oMockModel.requests[i].inputMaterialNumber,
					"Date": oMockModel.requests[i].inputDeliveryDate,
					"Price": oMockModel.requests[i].inputUnit,
					"Currency": oMockModel.requests[i].inputPrice
				}

				mockArr.push(row);
			}
			
			//First Sheet
			wb.SheetNames.push("Requests");
			let ws = XLSX.utils.json_to_sheet(mockArr);
			wb.Sheets["Requests"] = ws;
			let wscols = [{
				wch: 20
			}, {
				wch: 20
			}, {
				wch: 20
			}, {
				wch: 20
			}, {
				wch: 20
			}];
			ws['!cols'] = wscols;
			ws['!autofilter'] = {
				ref: "A1:E1"
			};
			
			//Second Sheet
			wb.SheetNames.push("Requests2");
			let ws2 = XLSX.utils.json_to_sheet(mockArr);
			wb.Sheets["Requests2"] = ws2;
			let wscols2 = [{
				wch: 20
			}, {
				wch: 20
			}, {
				wch: 20
			}, {
				wch: 20
			}, {
				wch: 20
			}];
			ws2['!cols'] = wscols2;
			ws2['!autofilter'] = {
				ref: "A1:E1"
			};

			const wbout = XLSX.write(wb, {
				bookType: 'xlsx',
				type: 'binary'
			});

			function s2ab(s) {
				let buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
				let view = new Uint8Array(buf); //create uint8array as viewer
				for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
				return buf;
			}
			saveAs(new Blob([s2ab(wbout)], {
				type: "application/octet-stream"
			}), 'Document.xlsx');
		},

		onExportExcelJs: function (oEvnet) {
			let oMockModel = this.getOwnerComponent().getModel("mock").getData();

			let workbook = new ExcelJS.Workbook();
			let worksheet = workbook.addWorksheet('Requests');
			let worksheet2 = workbook.addWorksheet('Requests2');

			// Define columns in the worksheet, these columns are identified using a key.
			worksheet.columns = [{
				header: 'Request',
				key: 'inputPRAmount',
				width: 20
			}, {
				header: 'Material',
				key: 'inputMaterialNumber',
				width: 20
			}, {
				header: 'Date',
				key: 'inputDeliveryDate',
				width: 20
			}, {
				header: 'Price',
				key: 'inputUnit',
				width: 20
			}, {
				header: 'Currency',
				key: 'inputPrice',
				width: 20
			}]

			// Define columns in the worksheet, these columns are identified using a key.
			worksheet2.columns = [{
				header: 'Request',
				key: 'inputPRAmount',
				width: 20
			}, {
				header: 'Material',
				key: 'inputMaterialNumber',
				width: 20
			}, {
				header: 'Date',
				key: 'inputDeliveryDate',
				width: 20
			}, {
				header: 'Price',
				key: 'inputUnit',
				width: 20
			}, {
				header: 'Currency',
				key: 'inputPrice',
				width: 20
			}]

			// Add rows from database to worksheet 
			for (const row of oMockModel.requests) {
				worksheet.addRow(row);
			};

			// Add rows from database to worksheet 
			for (const row of oMockModel.requests) {
				worksheet2.addRow(row);
			};

			// Add auto-filter on each column
			worksheet.autoFilter = 'A1:E1';

			// Add auto-filter on each column
			worksheet2.autoFilter = 'A1:E1';

			worksheet.eachRow((row, rowNumber) => {
				row.eachCell((cell, colNumber) => {
						if (rowNumber == 1) {
							// First set the background of header row
							cell.fill = {
								type: 'pattern',
								pattern: 'solid',
								fgColor: {
									argb: '778899'
								}
							};
						};
						// Set border of each cell 
						cell.border = {
							top: {
								style: 'thin'
							},
							left: {
								style: 'thin'
							},
							bottom: {
								style: 'thin'
							},
							right: {
								style: 'thin'
							}
						};
					})
					//Commit the changed row to the stream
				row.commit();
			});

			worksheet2.eachRow((row, rowNumber) => {
				row.eachCell((cell, colNumber) => {
						if (rowNumber == 1) {
							// First set the background of header row
							cell.fill = {
								type: 'pattern',
								pattern: 'solid',
								fgColor: {
									argb: 'f5b914'
								}
							};
						};
						// Set border of each cell 
						cell.border = {
							top: {
								style: 'thin'
							},
							left: {
								style: 'thin'
							},
							bottom: {
								style: 'thin'
							},
							right: {
								style: 'thin'
							}
						};
					})
					//Commit the changed row to the stream
				row.commit();
			});

			//Process 'Price' column for conditioning 
			const price = worksheet.getColumn('inputUnit');
			// Iterate over all current cells in this column
			price.eachCell((cell, rowNumber) => {
				// If the balance due is 400 or more, highlight it with gradient color 
				if (cell.value >= 400) {
					cell.fill = {
						type: 'gradient',
						gradient: 'angle',
						degree: 0,
						stops: [{
							position: 0,
							color: {
								argb: '4682B4'
							}
						}, {
							position: 0.5,
							color: {
								argb: '4682B4'
							}
						}, {
							position: 1,
							color: {
								argb: '4682B4'
							}
						}]
					};
				};
			});

			//Process 'Price' column for conditioning 
			const price2 = worksheet2.getColumn('inputUnit');
			// Iterate over all current cells in this column
			price2.eachCell((cell, rowNumber) => {
				// If the balance due is 400 or more, highlight it with gradient color 
				if (cell.value >= 400) {
					cell.fill = {
						type: 'gradient',
						gradient: 'angle',
						degree: 0,
						stops: [{
							position: 0,
							color: {
								argb: 'ffffff'
							}
						}, {
							position: 0.5,
							color: {
								argb: 'cc8188'
							}
						}, {
							position: 1,
							color: {
								argb: 'fa071e'
							}
						}]
					};
				};
			});

			this.onExport(workbook);
		},

		onExport: function (workbook) {
			workbook.xlsx.writeBuffer().then(function (buffer) {
				// done
				const blob = new Blob([buffer], {
					type: "applicationi/xlsx"
				});
				saveAs(blob, "Requests.xlsx");
			});
		}
	});
});