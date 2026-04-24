const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

const DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday'];

function localDateStr(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

function fmt(date) {
    return date.toLocaleDateString('en-AU', { day:'2-digit', month:'2-digit', year:'numeric' });
}

function toMins(t) {
    if (!t) return null;
    const [h, m] = t.split(':').map(Number);
    return h * 60 + m;
}

function minsToStr(mins) {
    if (mins === null || mins < 0) return null;
    return `${Math.floor(mins / 60)}h ${(mins % 60).toString().padStart(2, '0')}m`;
}

function minsToDecimal(mins) {
    if (mins === null) return null;
    return Math.round(mins / 60 * 100) / 100;
}

function calcDay(d) {
    const s = toMins(d.startTime), e = toMins(d.endTime);
    if (s === null || e === null || e <= s) return null;
    const ls = toMins(d.lunchStart), le = toMins(d.lunchEnd);
    let worked = e - s;
    if (ls !== null && le !== null && le > ls && ls >= s && le <= e) worked -= (le - ls);
    return Math.max(worked, 0);
}

const dataPath = path.join(__dirname, '..', 'data', 'timedata.json');
if (!fs.existsSync(dataPath)) {
    console.log('No data file found — skipping export.');
    process.exit(0);
}

const rawData = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
const allKeys = Object.keys(rawData).filter(k => k.startsWith('lfc_')).sort();

if (!allKeys.length) {
    console.log('No time data to export.');
    process.exit(0);
}

const detailRows  = [];
const summaryRows = [];

allKeys.forEach(sk => {
    const weekData = rawData[sk];
    const [yr, mo, dy] = sk.replace('lfc_', '').split('-').map(Number);
    const monday = new Date(yr, mo - 1, dy);
    let weekMins = 0;

    DAYS.forEach((dayName, i) => {
        const dayDate = new Date(monday);
        dayDate.setDate(dayDate.getDate() + i);
        const dateKey = localDateStr(dayDate);
        const d = weekData[dateKey] || {};

        const ls = toMins(d.lunchStart), le = toMins(d.lunchEnd);
        const s  = toMins(d.startTime),  e  = toMins(d.endTime);
        const lunchValid = ls !== null && le !== null && le > ls && s !== null && ls >= s && e !== null && le <= e;
        const lm = lunchValid ? le - ls : null;

        const dayMins = calcDay(d);
        if (dayMins !== null) weekMins += dayMins;

        detailRows.push({
            'Week Starting':  fmt(monday),
            'Day':            dayName,
            'Date':           fmt(dayDate),
            'Start Time':     d.startTime  || '',
            'Lunch Start':    d.lunchStart || '',
            'Lunch End':      d.lunchEnd   || '',
            'Lunch Duration': lm !== null ? minsToStr(lm) : '',
            'End Time':       d.endTime    || '',
            'Hours Worked':   dayMins !== null ? minsToStr(dayMins)    : '',
            'Hours (Decimal)':dayMins !== null ? minsToDecimal(dayMins): '',
            'Notes':          d.notes || ''
        });
    });

    const friday = new Date(monday);
    friday.setDate(friday.getDate() + 4);
    summaryRows.push({
        'Week Starting':        fmt(monday),
        'Week Ending':          fmt(friday),
        'Total Hours':          minsToStr(weekMins),
        'Total Hours (Decimal)':minsToDecimal(weekMins),
        'Days Logged':          Object.values(weekData).filter(d => calcDay(d) !== null).length
    });
});

const wb = XLSX.utils.book_new();

const ws1 = XLSX.utils.json_to_sheet(detailRows);
ws1['!cols'] = [
    {wch:16},{wch:12},{wch:12},{wch:12},
    {wch:12},{wch:12},{wch:16},{wch:12},
    {wch:14},{wch:18},{wch:35}
];
XLSX.utils.book_append_sheet(wb, ws1, 'Time Records');

const ws2 = XLSX.utils.json_to_sheet(summaryRows);
ws2['!cols'] = [{wch:16},{wch:14},{wch:14},{wch:22},{wch:14}];
XLSX.utils.book_append_sheet(wb, ws2, 'Weekly Summary');

const exportsDir = path.join(__dirname, '..', 'exports');
if (!fs.existsSync(exportsDir)) fs.mkdirSync(exportsDir, { recursive: true });

const filename = `TimeTracker_${localDateStr(new Date())}.xlsx`;
XLSX.writeFile(wb, path.join(exportsDir, filename));
console.log(`Exported: exports/${filename}`);
