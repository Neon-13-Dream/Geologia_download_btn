import XlsxPopulate from 'xlsx-populate/browser/xlsx-populate'
import {
	uName, uType, otherValue, downloadFile
} from '../../config'


export const exportHandler = async (records: any, name: string, availableШndex: number = -1) => {
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

	if (name === "Excel ( Кунлик )") {
		records = { [`${cDate.getFullYear()} йил`]: records[`${cDate.getFullYear()} йил`] }
	}

	const workbook = await XlsxPopulate.fromBlankAsync()
	const sheet = workbook.sheet(0)

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
		wrapText: true,
		fill: 'B4C6E7'
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
			fontSize: 16,
			fill: 'ffffff'
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
	sheet.column(startXcell).width(45)
	sheet.range(startYcell, startXcell, startYcell + 1, startXcell,)
		.merged(true)
		.value(availableШndex < 0 ? "Худуд номи" : uName[availableШndex])
		.style(uTypetyle)

	sheet.range(startYcell, 5, startYcell, 7)
		.merged(true)
		.value("Шу жумладан")
		.style(uTypetyle)

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








	if (availableШndex < 0) {
		Object.keys(records).forEach((yearKey: any) => {
			records[yearKey].forEach((montKey: any, montIndex: number) => {
				if (montKey.length != 0) {
					uName.forEach((type: any, index: number) => {
						sheet.row(startYcell + index + 1).height(20)
						sheet.cell(startYcell + index + 1, startXcell)
							.value("  " + type + (isJami(index) ? `${Object.keys(records).length !== 1 ? ' '+yearKey : ''} (${montIndex + 1}-мониторинг)` : ''))
							.style({
								...mainStyle,
								horizontalAlignment: isJami(index) ? "center" : "left",
								fontSize: 12,
								fill: isJami(index) ? 'FFFF00' : index & 1 ? 'DDEBF7' : 'FFFFFF',
								bold: isJami(index)
							})
					})

					montKey.forEach((rows: any, yIndex: number) => {
						rows.forEach((item: any, xIndex: number) => {
							sheet.cell(yIndex + startYcell + 1, xIndex + startXcell + 1)
								.value(TestFormat(item.count, item.sum, item.count ? 0 : -1))
								.style({
									border: true,
									bold: isJami(yIndex),
									horizontalAlignment: "center",
									verticalAlignment: "center",
									fontSize: 14,
									fill: isJami(yIndex) ? 'FFFF00' : yIndex & 1 ? 'DDEBF7' : 'FFFFFF',
								})
						})
					})
					startYcell += uName.length
				}
			})
		})
	}
	else {
		Object.keys(records).forEach((yearKey: any, yearIndex: number) => {
			records[yearKey].forEach((montKey: any, montIndex: number) => {
				if (montKey.length != 0) {
					sheet.row(startYcell + 1).height(20)
					sheet.cell(startYcell + 1, startXcell)
						.value(`  ${yearKey} (${montIndex + 1}-мониторинг)`)
						.style({
							...mainStyle,
							horizontalAlignment: "center",
							fontSize: 12,
							fill: yearIndex & 1 ? 'DDEBF7' : 'FFFFFF',
							bold: true
						})

					montKey[availableШndex].forEach((item: any, xIndex: number) => {
						sheet.cell(startYcell + 1, xIndex + startXcell + 1)
							.value(TestFormat(item.count, item.sum, item.count ? 0 : -1))
							.style({
								border: true,
								horizontalAlignment: "center",
								verticalAlignment: "center",
								fontSize: 14,
								fill: yearIndex & 1 ? 'DDEBF7' : 'FFFFFF',
							})
					})

					startYcell += 1
				}
			})
		})
	}










	const blob = await workbook.outputAsync()
	downloadFile(blob, `Geologiya ${name}.xlsx`)
}
