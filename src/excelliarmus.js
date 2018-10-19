// Excelliarmus v1.0 - Your tiny excel generator.
// Michele LEO 2018
// github.com/Leodau/Excelliarmus

function Excelliarmus() {
	this.sheets = [];
	this.styleSheet = new ExcelliarmusStyleSheet();
	this.Templates = {
		relations: "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>",
		workbookRelations: "<Relationship Id='rId2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/>",
		baseWorkbook: "<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/>",
		baseWorkkbookRelations: "<?xml version='1.0' ?><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>",
		workbook: "<?xml version='1.0' standalone='yes'?><workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets>",
		contentTypes: "<?xml version='1.0' encoding='UTF-8' standalone='yes' ?><Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default ContentType='application/xml' Extension='xml'/> <Default ContentType='application/vnd.openxmlformats-package.relationships+xml' Extension='rels'/> <Override ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' PartName='/xl/workbook.xml'/><Override ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' PartName='/xl/styles.xml'/>",
		styleRelation: "<Relationship Id='rId2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/>"
	};
	this.Options = {
		fileName: "Excelliarmus",
	};
	this.__selectedSheet = 0;

	function ExcelliarmusSheet(id, name) {
		this.worksheetTemplate = '<?xml version="1.0" ?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"> [sheetViews] [columns] <sheetData> [rows] </sheetData> [extras] </worksheet>';
		this.id = (id + 1);
		this.reelId = "rId" + (3 + id);
		this.name = name || "Sheet " + id;
		this.conditionals = [];
		this.data = [];
		this.yMax = 0;
		this.freezed = {xSplit: 0, ySplit: 0, topLeftCell: "Z9"};
		this.selected = false;
		this.extras = [];

		this.__push = (x, y, cell) => {
			if (isNaN(cell.value) && cell.value.charAt(0) == '=')
				cell.isFormula = true;
			if (!this.data[x])
				this.data[x] = [];
			cell.x = x - 1;
			cell.y = y;
			this.data[x][y] = cell;
			this.yMax = (y > this.yMax ? y : this.yMax);
		};
		this.__insert = (request, styleSheet) => {
			let data = {};

			if (!Excelliarmus.isInteger(request.value) && !request.value) throw "ExcelliarmusJS: (Insert) no 'value' found.";
			if (request.style) {
				if (Excelliarmus.isInteger(request.style)) {
					if (!styleSheet.styles[request.style - 1]) throw "ExcelliarmusJS: (Insert) requested style not found.";
				} else if (Excelliarmus.isObject(request.style)) {
					request.style = styleSheet.addStyle(request.style);
				} else throw "ExcelliarmusJS: (Insert) bad/wrong style.";
			}
			data = {value: (request.value || (Excelliarmus.isInteger(request.value) ? "0" : " "))};
			if (request.style !== undefined) data.style = request.style;
			if (request.column && request.row) {
				this.__push(request.column, request.row, data);
			} else if (request.column) {
				// Column insert not yet supported.
			} else if (request.row) {
				// Row insert not yet supported.
			}
		};
		this.__insertBlock = (request, conditionalObj, conditionalStyleSheet) => {
			let offset = this.yMax + 1;

			if (conditionalObj && conditionalStyleSheet) {
				request.forEach((value, col) => {
					this.__insert({
						column: col + (conditionalObj.startCol || 1),
						row: (conditionalObj.row || offset),
						value: (value || (Excelliarmus.isInteger(value) ? "0" : " ")),
						style: conditionalObj.style},
						conditionalStyleSheet
					);
				});
			} else {
				request.forEach((value, col) => {
					this.__insert({
						column: col + 1,
						row: offset,
						value: (value || (Excelliarmus.isInteger(value) ? "0" : " "))
					}, true);
				});
			}
		};
		this.__freezePane = (x, y) => {
			this.freezed.xSplit = (x || 0);
			this.freezed.ySplit = (y || 0);
			this.freezed.topLeftCell = Excelliarmus.intToColumn(this.freezed.xSplit) + (this.freezed.ySplit + 1);
		};
		this.__setConditionalColoring = (request) => {
			this.extras.push({
				column: Excelliarmus.intToColumn(request.column - 1),
				start: (request.start || 1),
				end: (request.end || 999999),
				specific: (request.type === "specific" ? true : false)
			});
		};
		this.__ExportExtras = () => {
			let output = "";

			if (!this.extras || !this.extras.length) return (output);
			this.extras.forEach((extra, i) => {
				output += "<conditionalFormatting sqref='" + (extra.column + extra.start.toString()) + ":" + (extra.column + extra.end.toString()) + "'>";
				output += "<cfRule type='colorScale' priority='" + (35 + i) + "'>";
				output += "<colorScale><cfvo type='min'/>";
				if (extra.specific) {
					output += "<cfvo type='percentile' val='1'/><cfvo type='max'/>";
					output += "<color rgb='FFFFFFFF'/><color rgb='FFFFFFFF'/><color rgb='FFF8696B'/>";
				} else {
					output += "<cfvo type='percentile' val='25'/><cfvo type='max'/>";
					output += "<color rgb='FF63BE7B'/><color rgb='FFFFEB84'/><color rgb='FFF8696B'/>";
				}
				output += "</colorScale></cfRule></conditionalFormatting>";
			});
			return (output);
		};
		this.__ExportPane = () => {
			let output = "<pane";

			output += (this.freezed.xSplit ? (" xSplit='" + this.freezed.xSplit + "'") : "");
			output += (this.freezed.ySplit ? (" ySplit='" + this.freezed.ySplit + "'") : "");
			output += " topLeftCell='" + this.freezed.topLeftCell + "'";
			output += " activePane='bottomRight' state='frozen'/>";
			return (output);
		};
		this.__ExportViews = () => {
			let output = "<sheetViews>";

			output += "<sheetView" + (this.selected ? " tabSelected='1'" : "") + " workbookViewId='0'";
			output += ((this.freezed.xSplit !== 0 || this.freezed.ySplit !== 0) ? (">" + this.__ExportPane() + "</sheetView>") : "/>");
			output += "</sheetViews>";
			return (output);
		};
		this.__ExportCell = (data) => {
			let output = "<c";

			output += " r='" + Excelliarmus.intToColumn(data.x) + data.y + "'";
			output += ((data.style) ? (" s='" + data.style + "'") : "");
			if (data.isFormula) {
				output += ">"
				output += ("<f>" + data.value.substring(1) + "</f>");
			} else if (isNaN(data.value)) {
				output += (" t='inlineStr'>");
				output += "<is><t>" + data.value;
				output += "</t></is>";
			} else {
				output += ">"
				output += "<v>" + data.value + "</v>";
			}
			output += "</c>";
			return (output);
		};
		this.__ExportRows = () => {
			let output = "";
			let rows = [];

			this.data.forEach((column, x) => {
				column.forEach((row, y) => {
					if (!rows[y]) rows[y] = [];
					rows[y].push(this.data[x][y]);
				});
			});
			rows.forEach((row, i) => {
				output += ("<row r='" + (i) + "'>");
				row.forEach((cell, j) => {
					output += this.__ExportCell(cell);
				})
				output += "</row>";
			});
			return (output);
		};
		this.__ExportColumns = () => {
			let output = "<cols>";
			
			if (!this.data.length) return;
			this.data.forEach((column, i) => {
				output += "<col min='" + (i) + "' max='" + (i) + "'";
				if (!column.width) {
					output += " width='25' bestFit='1' customWidth='1' ";
				} else {
					output += " width='" + column.width + "' customWidth='1' ";
				}
				if (column.style)
					output += " style='" + column.style + "'";
				output += "/>";
			});
			return (output + "</cols>");
		};
		this.__ExportData = () => {
			let file = this.worksheetTemplate;

			file = file.replace("[sheetViews]", this.__ExportViews());
			file = file.replace("[columns]", this.__ExportColumns());
			file = file.replace("[rows]", this.__ExportRows());
			file = file.replace("[extras]", this.__ExportExtras());
			return (file);
		};
	}
	function ExcelliarmusStyleSheet() {
		this.formats = ["General","0","0.00","#,##0","#,##0.00","0%","0.00%","0.00E+00","# ?/?","# ??/??","mm-dd-yy","d-mmm-yy","d-mmm","mmm-yy","h:mm AM/PM","h:mm:ss AM/PM","h:mm","h:mm:ss","m/d/yy h:mm","[$-404]e/m/d","m/d/yy","[$-404]e/m/d","#,##0 (#,##0)","#,##0 [Red](#,##0)","#,##0.00(#,##0.00)","#,##0.00[Red](#,##0.00)",'_("$"* #,##0.00_)_("$"* (#,##0.00)_("$"* "-"??_)_(@_)',"mm:ss","[h]:mm:ss","mmss.0","##0.0E+0","@","[$-404]e/m/d","[$-404]e/m/d","t0","t0.00","t#,##0","t#,##0.00","t0%","t0.00%","t# ?/?","t# ??/??"];
		this.Templates = {
			header: "<?xml version='1.0' encoding='utf-8'?><styleSheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' mc:Ignorable='x14ac' xmlns:x14ac='http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'>",
			cellStyleXfs: "<cellStyleXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0'/></cellStyleXfs>",
			cellStylesEnd: "<cellStyles count='1'><cellStyle name='Normal' xfId='0' builtinId='0'/></cellStyles><dxfs count='0'/>",
		};
		this.styles = [];
		this.fills = [];
		this.fonts = [];
		this.borders = [];
		this.supportedStyles = {
			fontStyles: ["bold", "italic", "underlined"],
			textAlign: ["left", "center", "right", "top", "middle"],
			borderStyles: ["thin", "medium", "double", "dotted"],
			formats: {percentage: 9, currency: 5}
		};

		__install = (list, element) => {
			let pos;

			if (element === undefined || !Object.keys(element).length) return (-1);
			pos = list.indexOf(element);
			if (pos != -1) return (pos);
			list.push(element);
			return (list.length - 1);
		};
		__readBorder = (str) => {
			let color;
			let type;
			let result = {};

			if (!str || str.length < 4 || str.includes("none")) return (undefined);
			type = Excelliarmus.wichExists(this.supportedStyles.borderStyles, str.split(" "));
			color = Excelliarmus.any(Excelliarmus.isColor, str.split(" "));
			if (!type && !color) return (undefined);
			result.type = (type || "thin");
			result.color = (color || "#000000");
			return (result);
		};
		this.__InstallFont = (style) => {
			let newFont = {};

			if (!style.fontSize && !style.fontFamily && !style.fontStyle && !style.fontColor) return (-1);
			if (style.fontSize !== undefined && parseInt(style.fontSize, 10))
				newFont.fontSize = style.fontSize;
			if (style.fontStyle !== undefined && this.supportedStyles.fontStyles.includes(style.fontStyle))
				newFont.fontStyle = style.fontStyle;
			if (style.fontFamily !== undefined)
				newFont.fontFamily = style.fontFamily;
			if (style.fontColor != undefined && style.fontColor.charAt(0) === "#" && style.fontColor.length === 7)
				newFont.fontColor = style.fontColor.toString().substring(1).toUpperCase();
			return (__install(this.fonts, newFont));
		};
		this.__InstallBorder = (style) => {
			let element = {};

			if (!style.borders && !style.border && !style.borderTop && !style.borderRight && !style.borderBottom && !style.borderLeft) return (-1);
			if (style.border) {
				element.top = __readBorder(style.border);
				element.right = element.bottom = element.left = element.top;
			}
			if (style.borders) {
				let splitted = style.borders.split(",");
				return (this.__InstallBorder({borderTop: splitted[0], borderRight: splitted[1], borderBottom: splitted[2], borderLeft: splitted[3]}));
			}
			if (style.borderTop) element.top = __readBorder(style.borderTop);
			if (style.borderRight) element.right = __readBorder(style.borderRight);
			if (style.borderBottom) element.bottom = __readBorder(style.borderBottom);
			if (style.borderLeft) element.left = __readBorder(style.borderLeft);
			return ((Object.keys(element).length ? __install(this.borders, element) : -1));
		}
		this.__GenerateFormats = () => {
			let output = "<numFmts count='" + (this.formats.length) + "'>"

			this.formats.forEach((format, i) => {
				output += "<numFmt numFmtId='" + (i) + "' formatCode='" + format + "'/>";
			});
			output += "</numFmts>";
			return (output);
		};
		this.__GenerateFills = () => {
			let output = "<fills count='"+ (this.fills.length + 2) +"'><fill><patternFill patternType='none'/></fill><fill><patternFill patternType='gray125'/></fill>";

			this.fills.forEach((fill, i) => {
				output += "<fill><patternFill patternType='solid'><fgColor rgb='FF" + fill + "' /><bgColor indexed='64'/></patternFill></fill>";
			});
			output += "</fills>";
			return (output);
		};
		this.__GenerateFonts = () => {
			let output = "<fonts count='"+ (this.fonts.length + 1) + "' x14ac:knownFonts='1' xmlns:x14ac='http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'>";

			output += "<font><sz val='12'/><name val='Calibri Light'/></font>";
			this.fonts.forEach((font, i) => {
				output += "<font>";
				output += "<sz val='" + (font.fontSize || "12") + "'/>";
				output += (font.fontStyle ? ("<" + font.fontStyle[0] + "/>") : "");
				output += (font.fontFamily ? ("<name val='" + font.fontFamily.toString() + "'/>") : "");
				output += (font.fontColor ? ("<color rgb='FF" + font.fontColor + "'/>") : "");
				output += "</font>";
			});
			output += "</fonts>";
			return (output);
		};
		this.__GenerateBorders = () => {
			let output = "<borders count='"+ (this.borders.length + 1) + "'><border><left/><right/><top/><bottom/><diagonal/></border>";

			this.borders.forEach((border, i) => {
				output += "<border>";
				output += (border.left ? "<left style='" + border.left.type + "'><color rgb='FF" + border.left.color.toString().substring(1).toUpperCase() + "'/></left>" : "<left/>");
				output += (border.right ? "<right style='" + border.right.type + "'><color rgb='FF" + border.right.color.toString().substring(1).toUpperCase() + "'/></right>" : "<right/>");
				output += (border.top ? "<top style='" + border.top.type + "'><color rgb='FF" + border.top.color.toString().substring(1).toUpperCase() + "'/></top>" : "<top/>");
				output += (border.bottom ? "<bottom style='" + border.bottom.type + "'><color rgb='FF" + border.bottom.color.toString().substring(1).toUpperCase() + "'/></bottom>" : "<bottom/>");
				output += "<diagonal/></border>";
			});
			output += "</borders>";
			return (output);
		};
		this.__GenerateXfs = () => {
			let output = "<cellXfs count='" + (this.styles.length + 1) + "'><xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/>";

			this.styles.forEach((style, i) => {
				output += this.__ExportStyle(style);
			});
			output += "</cellXfs>";
			return (output);
		};
		this.__ExportStyle = (style) => {
			let output = "<xf numFmtId='" + style.format + "' fontId='" + style.font + "' fillId='" + style.fill + "' borderId='" + style.border + "' xfId='0'";

			if (style.fill) output += " applyFill='1'";
			if (style.border) output += " applyBorder='1'";
			if (style.horizontalAlign || style.verticalAlign) {
				output += " applyAlignment='1'";
				output += ">";
				output += "<alignment";
				output += (style.horizontalAlign ? " horizontal='" + style.horizontalAlign + "'" : "");
				output += (style.verticalAlign ? " vertical='" + style.verticalAlign + "'" : "");
				output += "/>";
			} else {
				output += ">";
			}
			output += "</xf>";
			return (output);
		};
		this.__Generate = () => {
			let output = this.Templates.header;

			output += this.__GenerateFonts();
			output += this.__GenerateFills();
			output += this.__GenerateBorders();
			output += this.Templates.cellStyleXfs;
			output += this.__GenerateXfs();
			output += this.Templates.cellStylesEnd;
			output += "</styleSheet>";
			return (output);
		};
		this.addStyle = (wanted) => {
			let newStyle = {};
			let fontId;
			let borderId;
			let horizontal;
			let vertical;

			if (wanted.backgroundColor !== undefined && wanted.backgroundColor.charAt(0) == "#")
				newStyle.fill = 2 + __install(this.fills, wanted.backgroundColor.toString().substring(1).toUpperCase());
			if (wanted.textAlign !== undefined && (horizontal = Excelliarmus.wichExists(this.supportedStyles.textAlign.slice(0, 3), wanted.textAlign.split(" ")) ))
				newStyle.horizontalAlign = horizontal;
			if (wanted.textAlign !== undefined && (vertical = Excelliarmus.wichExists(this.supportedStyles.textAlign.slice(-2), wanted.textAlign.split(" ")) ))
				newStyle.verticalAlign = vertical.replace("middle", "center");
			borderId = this.__InstallBorder(wanted);
			fontId = this.__InstallFont(wanted);
			if (fontId > -1)
				newStyle.font = fontId + 1;
			if (borderId > -1)
				newStyle.border = borderId + 1;
			if (wanted.format && this.supportedStyles.formats.hasOwnProperty(wanted.format))
				newStyle.format = this.supportedStyles.formats[wanted.format];
			if (!Object.keys(newStyle).length)
				throw "ExcelliarmusJS: (AddStyle) : Empty style or Unsupported properties.";
			this.styles.push(newStyle);
			return (this.styles.length);
		};
	}

	__GenerateSheetWorkbooks = (sheets) => {
		let output = this.Templates.workbook;

		sheets.forEach(element => {
			output += "<sheet state='visible' name='" + element.name + "' sheetId='" + element.id + "' r:id='" + element.reelId + "' />";
		});
		output += "</sheets><calcPr/></workbook>";
		return (output);
	};
	__GenerateWorkbookRelations = (sheets) => {
		let output = this.Templates.baseWorkkbookRelations;

		output += this.Templates.styleRelation;
		sheets.forEach((element, i) => {
			output += '<Relationship Id="' + element.reelId + '" Target="worksheets/sheet' + (i + 1) + '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>';
		});
		output += "</Relationships>";
		return (output);
	};
	__GenerateContentTypes = (sheets) => {
		let output = this.Templates.contentTypes;

		sheets.forEach((element, i) => {
			output += "<Override ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' PartName='/xl/worksheets/sheet" + (i + 1) + ".xml'/>";
		});
		return (output + "</Types>")
	};
	__GenerateSheetXml = (sheets, folder) => {
		sheets.forEach((element, i) => {
			folder.file("worksheets/sheet" + (i + 1) + ".xml", element.__ExportData());
		});
	};
	this.setActiveSheet = (id) => {
		if (!id) return;
		if (!this.sheets[id]) throw "ExcelliarmusJS: (setActiveSheet) Requested sheet id doesn't exist.";
		this.__selectedSheet = id;
		this.sheets.forEach((sheet) => {sheet.selected = false;});
		this.sheets[id].selected = true;
	};
	this.freeze = (request) => {
		if (!request || (!request.column && !request.row)) throw "ExcelliarmusJS: (Freeze) Bad request. (Sheet, Column ||&& Row) are required.";
		if (request.sheet && !this.sheets[request.sheet]) throw "ExcelliarmusJS: (Freeze) Requested sheet id doesn't exist.";
		this.sheets[(request.sheet || 0)].__freezePane(request.column, request.row);
	};
	this.setConditional = (request) => {
		if (!request || !request.column || (request.column < 1)) throw "ExcelliarmusJS: (Conditional Formatting) Bad request. (Atleast a Column is required).";
		if (request.start && request.start < 1) throw "ExcelliarmusJS: (Conditional Formatting) Start has to be positive.";
		if ((request.end && request.end < 1) || ((request.end && request.start) && (request.end <= request.start))) throw "ExcelliarmusJS: (Conditional Formatting) End has to be positive and greater than start.";
		if (!this.sheets[(request.sheet || 0)]) throw "ExcelliarmusJS: (Conditional Formatting) Bad request, no sheet found.";
		this.sheets[(request.sheet || 0)].__setConditionalColoring(request);
	};
	this.addSheet = (name) => {
		this.sheets.push(new ExcelliarmusSheet(this.sheets.length, name));
		if (this.sheets.length === 1) this.sheets[0].selected = true;
		return (this.sheets.length - 1);
	};
	this.addStyle = (request) => {
		if (!request) throw "ExcelliarmusJS: (InsertStyle) Bad request.";
		return (this.styleSheet.addStyle(request));
	};
	this.insert = (request, selector) => {
		let sheet;

		if (!this.sheets.length)
			throw "ExcelliarmusJS: (Insert) You need to create a sheet first.";
		if (Array.isArray(request)) {
			if (selector) {
				if (Excelliarmus.isInteger(selector)) {
					if (!this.sheets[selector]) throw "ExcelliarmusJS: (Insert) bad sheet id.";
					return (this.sheets[selector].__insertBlock(request));
				} else if (Excelliarmus.isObject(selector)) {
					return (this.sheets[selector.sheet || 0].__insertBlock(request, selector, this.styleSheet));
				}
				throw "ExcelliarmusJS: (Insert as block) bad usage.";
			}
			return (this.sheets[0].__insertBlock(request));
		}
		if (request.sheet && (!Excelliarmus.isInteger(request.sheet) || !this.sheets[request.sheet]))
			throw "ExcelliarmusJS: (Insert) bad sheet id.";
		sheet = this.sheets[request.sheet || 0];
		if ((!request.column && !request.row))
			throw "ExcelliarmusJS: (Insert) bad col or row: ["+ request.column + ":" + request.row +"]";
		if ((request.column && !Excelliarmus.isInteger(request.column)) || (request.row && !Excelliarmus.isInteger(request.row)))
			throw "ExcelliarmusJS: (Insert) col and row must be integers.";
		sheet.__insert(request, this.styleSheet);
	};
	this.export = (fileName) => {
		let name = (fileName || this.Options.fileName);
		let zip = new JSZip();
		let folder = zip.folder("xl");

		zip.file("_rels/.rels", this.Templates.relations + this.Templates.baseWorkbook + "</Relationships>");
		folder.file("workbook.xml", __GenerateSheetWorkbooks(this.sheets));
		folder.file("styles.xml", this.styleSheet.__Generate());
		folder.file("_rels/workbook.xml.rels", __GenerateWorkbookRelations(this.sheets));
		zip.file("[Content_Types].xml", __GenerateContentTypes(this.sheets));
		__GenerateSheetXml(this.sheets, folder);
		zip.generateAsync({
			type: "blob",
			mimeType: "application/vnd.ms-excel"
			}).then(function (content) {
				saveAs(content, name + ".xlsx");
		}); 
	};
}

Excelliarmus.intToColumn = (n) => {
	let rest = Math.floor(n / 26) - 1;
	let s = (rest > -1 ? Excelliarmus.intToColumn(rest) : '');

	return (s + "ABCDEFGHIJKLMNOPQRSTUVWXYZ".charAt(n % 26));
};
Excelliarmus.isInteger = (n) => {
	return (typeof(n) ==='number' && (n%1) === 0);
};
Excelliarmus.isObject = (o) => {
	return (o !== undefined && o !== null && typeof(o) === 'object');
};
Excelliarmus.wichExists = (s, n) => {
	while (!s.includes(n[0]) && n.shift());
	return (n[0]);
};
Excelliarmus.isColor = (c) => {
	return ((c && c.length === 7 && c[0] === "#") ? true : false);
};
Excelliarmus.any = (func, n) => {
	while (!func(n[0]) && n.shift());
	return (n[0]);
};