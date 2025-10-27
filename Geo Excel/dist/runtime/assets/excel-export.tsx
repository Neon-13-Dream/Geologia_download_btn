import XlsxPopulate from 'xlsx-populate/browser/xlsx-populate'
import {
	uName, uType, otherValue, downloadFile
} from '../../config'


export const exportHandler = async (records: any, name: any, current: boolean = false) => {
	if (Object.keys(records).length === 0) {
		console.log("IsEmpty")
		return
	}
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
		return val === uName.indexOf("Жами")
	}

	if (current) {
		records = { [`${cDate.getFullYear()} йил`]: records[`${cDate.getFullYear()} йил`] }
	}

	const workbook = await XlsxPopulate.fromBlankAsync()
	const sheet = workbook.sheet(0)
	console.log(records)

	const mainStyle = {
		bold: true,
		border: true,
		horizontalAlignment: "center",
		verticalAlignment: "center",
		fill: 'F2F2F2'
	}
	const uTypetyle = {
		...mainStyle,
		fontSize: 14,
		wrapText: true
	}

	const startXcell = 2
	let startYcell = 2

	sheet.row(startYcell).height(72)
	sheet.range(startYcell, startXcell, startYcell, startXcell + uType.length)
		.merged(true)
		.value("Ўзбекистон Республикаси Президентининг 2024 йил 14 октябрдаги\nПФ-155-сон Фармони 1-иловаси 2-банди ижроси бўйича кунлик\nТАҲЛИЛИЙ ЖАДВАЛ")
		.style({
			...mainStyle,
			wrapText: true,
			fontSize: 16
		})
	startYcell += 1
	sheet.row(startYcell).height(15)
	sheet.range(startYcell, startXcell, startYcell, startXcell + uType.length)
		.merged(true)
		.value(`${cDate.toLocaleDateString()} йил холатига`)
		.style({
			...mainStyle,
			horizontalAlignment: 'right',
			fontSize: 12
		})
	startYcell += 1

	sheet.row(startYcell).height(30)
	sheet.row(startYcell + 1).height(55)
	sheet.column(startXcell).width(40)
	sheet.range(startYcell, startXcell, startYcell + 1, startXcell,)
		.merged(true)
		.value("Худуд номи")
		.style(uTypetyle)

	sheet.range(startYcell, 5, startYcell, 7)
		.merged(true)
		.value("Шу жумладан")
		.style({
			...mainStyle,
			fontSize: 14
		})

	uType.forEach((type: any, index: number) => {
		if (otherValue.includes(index)) {
			sheet.column(startXcell + index + 1).width(30)
			sheet.cell(startYcell + 1, startXcell + index + 1)
				.value(type)
				.style(uTypetyle)
		}
		else {
			sheet.column(startXcell + index + 1).width(35)
			sheet.range(startYcell, startXcell + index + 1, startYcell + 1, startXcell + index + 1)
				.merged(true)
				.value(type)
				.style(uTypetyle)
		}
	})
	startYcell += 1
	sheet.freezePanes(0, startYcell)
	Object.keys(records).forEach((key: any) => {
		uName.forEach((type: any, index: number) => {
			sheet.row(startYcell + index + 1).height(20)
			sheet.cell(startYcell + index + 1, startXcell)
				.value("  " + type + (isJami(index) && Object.keys(records).length !== 1 ? ` ${key}` : ''))
				.style({
					...mainStyle,
					horizontalAlignment: isJami(index) ? "center" : "left",
					fontSize: 12,
					fill: isJami(index) ? 'D8D8D8' : 'F2F2F2',
					bold: isJami(index)
				})
		})

		records[key].forEach((rows: any, yIndex: number) => {
			rows.forEach((item: any, xIndex: number) => {
				sheet.cell(yIndex + startYcell + 1, xIndex + startXcell + 1)
					.value(TestFormat(item.count, item.sum, item.count ? 0 : -1))
					.style({
						border: true,
						bold: isJami(yIndex),
						horizontalAlignment: "center",
						verticalAlignment: "center",
						fontSize: 14,
						fill: isJami(yIndex) ? 'D8D8D8' : 'FFFFFF',
					})
			})
		})
		startYcell += uName.length
	})

	return workbook.outputAsync()
	//downloadFile(blob, `Geologiya_Download-${name}.xlsx`)
}
