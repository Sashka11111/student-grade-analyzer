let workbookData = null
let lastResults = []
let fileName = ''

document.getElementById('current-year').textContent = new Date().getFullYear();

lucide.createIcons()

document.getElementById('file-input').addEventListener('change', function (e) {
  const file = e.target.files[0]
  if (!file) return
  fileName = file.name
  processFile(file)
})

const zone = document.getElementById('upload-zone')
if (zone) {
  zone.addEventListener('dragover', (e) => {
    e.preventDefault()
    zone.classList.add('drag-over')
  })
  zone.addEventListener('dragleave', () => {
    zone.classList.remove('drag-over')
  })
  zone.addEventListener('drop', (e) => {
    e.preventDefault()
    zone.classList.remove('drag-over')
    const dt = e.dataTransfer
    if (dt.files[0]) {
      fileName = dt.files[0].name
      processFile(dt.files[0])
    }
  })
}

function processFile(file) {
  showLoading(true)
  const reader = new FileReader()
  reader.onload = function (ev) {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'binary' })
      workbookData = wb
      showConfig(file.name, wb)
      showLoading(false)
    } catch (err) {
      showLoading(false)
      showError('Не вдалося прочитати файл: ' + err.message)
    }
  }
  reader.readAsBinaryString(file)
}

function showLoading(show) {
  const overlay = document.getElementById('loading-overlay')
  if (overlay) overlay.style.display = show ? 'flex' : 'none'
}

function downloadTemplate() {
  const wb = XLSX.utils.book_new()
  const subjectsCount = 48

  const header1 = ['№', 'ПІБ студента']
  for (let i = 1; i <= subjectsCount; i++) header1.push('Предмет ' + i)

  const header2 = ['1', '2']
  for (let i = 1; i <= subjectsCount; i++) header2.push((i + 2).toString())

  const ws_data = [header1, header2]

  for (let s = 1; s <= 20; s++) {
    const row = [s.toString(), '']
    for (let i = 1; i <= subjectsCount; i++) {
      row.push('')
    }
    ws_data.push(row)
  }

  const ws = XLSX.utils.aoa_to_sheet(ws_data)

  const colWidths = [{ wch: 5 }, { wch: 35 }]
  for (let i = 0; i < subjectsCount; i++) colWidths.push({ wch: 10 })
  ws['!cols'] = colWidths

  XLSX.utils.book_append_sheet(wb, ws, 'Відомість')
  XLSX.writeFile(wb, 'шаблон_відомості.xlsx')
}

function showConfig(name, wb) {
  hideError()
  const configZone = document.getElementById('config-zone')
  if (configZone) {
    configZone.style.display = 'block'
    document.getElementById('results-zone').style.display = 'none'
    document.getElementById('group-stats').style.display = 'none'
    const sheet = wb.Sheets[wb.SheetNames[0]]
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:Z100')
    document.getElementById('cfg-end').value = range.e.r + 1
    document.getElementById('cfg-subjects').value = 48

    configZone.scrollIntoView({ behavior: 'smooth' })
  }
}

function parseGrade(val) {
  if (val === null || val === undefined || val === '') return null
  let str = String(val).replace(',', '.').trim()
  let num = parseFloat(str)
  return isNaN(num) ? null : num
}

function getAdjustedPercentages(counts, total) {
  if (total <= 0) return counts.map(() => '0.0%')

  let percentages = counts.map((c) => (c / total) * 100)
  let rounded = percentages.map((p) => Math.floor(p * 10) / 10)
  let sumRounded = rounded.reduce((a, b) => a + b, 0)
  let diff = Math.round((100 - sumRounded) * 10) / 10

  if (diff > 0) {
    let remainders = percentages.map((p, i) => ({
      index: i,
      rem: (p * 10) % 1,
    }))
    remainders.sort((a, b) => b.rem - a.rem)

    for (let i = 0; i < Math.round(diff * 10); i++) {
      rounded[remainders[i % remainders.length].index] += 0.1
    }
  }

  return rounded.map((r) => r.toFixed(1) + '%')
}

function calculate() {
  hideError()
  if (!workbookData) return

  const startRow = parseInt(document.getElementById('cfg-start').value)
  const endRow = parseInt(document.getElementById('cfg-end').value)
  const kolPred = parseInt(document.getElementById('cfg-subjects').value)

  if (startRow >= endRow || kolPred < 1) {
    showError('Перевірте налаштування (діапазон рядків або кількість предметів)')
    return
  }

  const sheet = workbookData.Sheets[workbookData.SheetNames[0]]
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })

  const results = []
  let groupSum = 0
  let qualityStudents = 0
  let successStudents = 0

  for (let i = startRow - 1; i < Math.min(endRow, data.length); i++) {
    const row = data[i] || []
    const name = (row[1] || row[0] || 'Студент ' + (i - startRow + 2)).trim()
    if (!name && i > startRow + 10) break

    let a = 0,
      b = 0,
      c = 0,
      d = 0,
      e = 0,
      f = 0,
      fx = 0
    let studentSum = 0
    let studentCount = 0

    for (let j = 2; j < 2 + kolPred; j++) {
      const tm = parseGrade(row[j])
      if (tm === null) continue

      studentSum += tm
      studentCount++

      if (tm >= 89.5) a++
      else if (tm >= 81.5) b++
      else if (tm >= 73.5) c++
      else if (tm >= 63.5) d++
      else if (tm >= 59.5) e++
      else if (tm >= 34.5) fx++
      else if (tm > 0) f++
    }

    if (studentCount === 0) continue

    const avg = studentSum / studentCount
    groupSum += avg

    if (avg >= 73.5) qualityStudents++
    if (avg >= 59.5) successStudents++

    const ectsCounts = [a, b, c, d, e, f, fx]
    const ectsPcts = getAdjustedPercentages(ectsCounts, studentCount)

    const summaryCounts = [a, b + c, d + e, f + fx]
    const summaryPcts = getAdjustedPercentages(summaryCounts, studentCount)

    results.push({
      name,
      ects: ectsPcts,
      summary: summaryPcts,
      total: studentCount,
      gpa: avg.toFixed(1),
    })
  }

  if (results.length === 0) {
    showError('Не знайдено даних для розрахунку у вказаному діапазоні')
    return
  }

  const groupAvgEl = document.getElementById('group-avg')
  if (groupAvgEl) {
    groupAvgEl.textContent = (groupSum / results.length).toFixed(1)
    document.getElementById('group-quality').textContent = ((qualityStudents / results.length) * 100).toFixed(1) + '%'
    document.getElementById('group-success').textContent = ((successStudents / results.length) * 100).toFixed(1) + '%'
    document.getElementById('group-count').textContent = results.length
    document.getElementById('group-stats').style.display = 'block'
  }

  lastResults = results
  renderResults(results)
}

function renderResults(results) {
  const tbody = document.getElementById('results-body')
  if (!tbody) return
  tbody.innerHTML = ''

  results.forEach((r, idx) => {
    const tr = document.createElement('tr')
    tr.innerHTML = `
            <td class="text-center" style="color: var(--text-secondary);">${idx + 1}</td>
            <td style="font-weight: 500;">${r.name}</td>
            <td class="text-center ects-a">${r.ects[0]}</td>
            <td class="text-center ects-b">${r.ects[1]}</td>
            <td class="text-center ects-c">${r.ects[2]}</td>
            <td class="text-center ects-d">${r.ects[3]}</td>
            <td class="text-center ects-e">${r.ects[4]}</td>
            <td class="text-center ects-f">${r.ects[5]}</td>
            <td class="text-center ects-fx">${r.ects[6]}</td>
            <td class="text-center summary-val">${r.summary[0]}</td>
            <td class="text-center summary-val">${r.summary[1]}</td>
            <td class="text-center summary-val">${r.summary[2]}</td>
            <td class="text-center summary-val">${r.summary[3]}</td>
        `
    tbody.appendChild(tr)
  })

  const resultsZone = document.getElementById('results-zone')
  if (resultsZone) {
    resultsZone.style.display = 'block'
    resultsZone.scrollIntoView({ behavior: 'smooth' })
  }
}

function exportCSV() {
  if (!lastResults.length) return
  const header = [
    '№',
    'ПІБ студента',
    'A',
    'B',
    'C',
    'D',
    'E',
    'F',
    'FX',
    'Відмінно',
    'Добре',
    'Задовільно',
    'Незадовільно',
  ]
  const rows = lastResults.map((r, i) => {
    return [i + 1, r.name, ...r.ects, ...r.summary]
  })
  const csv = [header, ...rows].map((r) => r.map((v) => '"' + String(v).replace(/"/g, '""') + '"').join(',')).join('\n')
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = 'результати_оцінок.csv'
  a.click()
  URL.revokeObjectURL(url)
}

function showError(msg) {
  const el = document.getElementById('error-zone')
  if (el) {
    el.textContent = msg
    el.style.display = 'block'
    el.scrollIntoView({ behavior: 'smooth' })
  }
}

function hideError() {
  const el = document.getElementById('error-zone')
  if (el) el.style.display = 'none'
}
