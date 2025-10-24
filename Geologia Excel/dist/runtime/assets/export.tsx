import XlsxPopulate from 'xlsx-populate/browser/xlsx-populate'
import { saveAs } from 'file-saver'
import { types, names, other_value } from "../../config"

export default async function exportHandler(records: any, name: string, current: boolean = false, inZip: any = null) {
    if( Object.keys(records).length == 0 ) return
    function NumsFormat(num: any) {
        num = Math.round(num * 100) / 100
        num = num.toString().split('.')

        return num[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + (num[1] ? ',' + num[1] : '')
    }
    function TestFormat(count: number, area: number, type: number = 0) {
        switch (type) {
            case 0: return NumsFormat(count) + " ( " + NumsFormat(area) + " га )"
            default: return '-'
        }
    }

    const cDate = new Date()
    function isJami(val: number) {
        return val == names.indexOf("Жами")
    }

    if (current) {
        records = { [`${cDate.getFullYear()} йил`]: records[`${cDate.getFullYear()} йил`] }
    }

    const workbook = await XlsxPopulate.fromBlankAsync()
    const sheet = workbook.sheet(0)

    var mainStyle = {
        bold: true,
        border: true,
        horizontalAlignment: "center",
        verticalAlignment: "center",
        fill: 'F2F2F2'
    }
    var typeStyle = {
        ...mainStyle,
        fontSize: 14,
        wrapText: true
    }

    var start_x_cell = 2
    var start_y_cell = 2

    sheet.row(start_y_cell).height(72)
    sheet.range(start_y_cell, start_x_cell, start_y_cell, start_x_cell + types.length)
        .merged(true)
        .value("Ўзбекистон Республикаси Президентининг 2024 йил 14 октябрдаги\nПФ-155-сон Фармони 1-иловаси 2-банди ижроси бўйича кунлик\nТАҲЛИЛИЙ ЖАДВАЛ")
        .style({
            ...mainStyle,
            wrapText: true,
            fontSize: 16
        })
    start_y_cell += 1
    sheet.row(start_y_cell).height(15)
    sheet.range(start_y_cell, start_x_cell, start_y_cell, start_x_cell + types.length)
        .merged(true)
        .value(`${cDate.toLocaleDateString()} йил холатига`)
        .style({
            ...mainStyle,
            horizontalAlignment: 'right',
            fontSize: 12
        })
    start_y_cell += 1

    sheet.row(start_y_cell).height(30)
    sheet.row(start_y_cell + 1).height(55)
    sheet.column(start_x_cell).width(40)
    sheet.range(start_y_cell, start_x_cell, start_y_cell + 1, start_x_cell,)
        .merged(true)
        .value("Худуд номи")
        .style(typeStyle)

    sheet.range(start_y_cell, 5, start_y_cell, 7)
        .merged(true)
        .value("Шу жумладан")
        .style({
            ...mainStyle,
            fontSize: 14
        })

    types.forEach((type: any, index: number) => {
        if (other_value.includes(index)) {
            sheet.column(start_x_cell + index + 1).width(30)
            sheet.cell(start_y_cell + 1, start_x_cell + index + 1)
                .value(type)
                .style(typeStyle)
        }
        else {
            sheet.column(start_x_cell + index + 1).width(35)
            sheet.range(start_y_cell, start_x_cell + index + 1, start_y_cell + 1, start_x_cell + index + 1)
                .merged(true)
                .value(type)
                .style(typeStyle)
        }
    })
    start_y_cell += 1
    sheet.freezePanes(0, start_y_cell)
    Object.keys(records).forEach((key: any) => {
        names.forEach((type: any, index: number) => {
            sheet.row(start_y_cell + index + 1).height(20)
            sheet.cell(start_y_cell + index + 1, start_x_cell)
                .value("  " + type + (isJami(index) && Object.keys(records).length != 1 ? ` ${key}` : ''))
                .style({
                    ...mainStyle,
                    horizontalAlignment: isJami(index) ? "center" : "left",
                    fontSize: 12,
                    fill: isJami(index) ? 'D8D8D8' : 'F2F2F2',
                    bold: isJami(index)
                })
        })

        records[key].forEach((rows: any, y_index: number) => {
            rows.forEach((item: any, x_index: number) => {
                sheet.cell(y_index + start_y_cell + 1, x_index + start_x_cell + 1)
                    .value(TestFormat(item["count"], item["sum"], item["count"] ? 0 : -1))
                    .style({
                        border: true,
                        bold: isJami(y_index),
                        horizontalAlignment: "center",
                        verticalAlignment: "center",
                        fontSize: 14,
                        fill: isJami(y_index) ? 'D8D8D8' : 'FFFFFF',
                    })
            })
        })
        start_y_cell += names.length
    })

    var blob = await workbook.outputAsync()
    if (!inZip) {
        saveAs(blob, `Geologiya_Download-${name}.xlsx`)
    }
    else {
        inZip.file(`Geologiya_Download-${name}.xlsx`, blob);
    }
}