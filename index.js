const fs = require('fs')
const path = require('path')
const util = require('util')
const excelToJson = require('convert-excel-to-json')
const moment = require('moment')
const calendar = require('./googleCalendarAPI.js')

const PARAM = require('./param.json')

const DEBUG = true
const DEBUG_DIR = path.join(__dirname, 'DEBUG')

// =========================================================
// récupérer des infos dans la colonne "Observation"
// =========================================================

// le nombre de pax "25 pax"
function REGEX_PAX(el) {
    let res = /(^|\s)([0-9]+)\s*pax/gi.exec(el.N)
    return (res && res.length > 2) ? res[2]: 0
}

// le type de pension "HB" ou "BB" ou "BIVOUAC"
function REGEX_PENSION(el) {
    let res = /(^|\s)(HB|BB|[Bb]ivouacs?)($|\s)/g.exec(el.N)
    if (res && res.length > 2) {
        if (res == 'HB' || res == 'BB') return res;
        return 'BIVOUAC'
    }
    return 'BB'
}
// =========================================================
// =========================================================

const HEURE_PTIDEJ = 7;
const HEURE_DINER = 19;
const HEURE_ARRIVEE = 18;
const HEURE_DEPART = 9;

// COLOR IDS : https://eduardopereira.pt/wp-content/uploads/2012/06/google_calendar_api_event_color_chart.png
const COLORS = {
    breakfast: '1',
    dinner: '3',
    arrivee: '10',
    depart: '2'
}

let DOSSIER_SOURCE = PARAM.dossier_source;
let ind_repere = -1;
process.argv.forEach(function (val, index, array) {
	if (/[\\\/]index.js$/.test(val)) {
        ind_repere = index
    } else if (ind_repere > 0 && index == ind_repere + 1) {
        DOSSIER_SOURCE = val
    }
});
console.log('Recherche du fichier Excel dans le dossier source ' + DOSSIER_SOURCE)


async function main() {
    // 1. récupérer le fichier .xlsx? le plus récent du dossier source
    let mostrecent_xlsx_path = getMostRecentFile(DOSSIER_SOURCE, /\.xlsx?$/i)
    if (mostrecent_xlsx_path == '') throw 'Pas de fichier excel trouvé dans le dossier source :('
    console.log('Fichier excel trouvé : ' + mostrecent_xlsx_path)
    if (DEBUG) fs.writeFile(path.join(DEBUG_DIR, '1_most_recent_src_file.txt'), mostrecent_xlsx_path, 'utf8', _ => 1)
    // 2. on le transforme en JSON
    let data = getData(mostrecent_xlsx_path)
    if (data.error) throw data;
    console.log('Il y a ' + data.length + ' dossiers')
    if (DEBUG) fs.writeFile(path.join(DEBUG_DIR, '2_xls2data.json'), JSON.stringify(data, null, '\t'), 'utf8', _ => 1)
    // 3. on le transforme en JSON pour Google Calendar
    let dataGoogle = data2GoogleCalendarFormat(data)
    console.log('Il y a ' + dataGoogle.length + ' événements Google')
    if (DEBUG) fs.writeFile(path.join(DEBUG_DIR, '3_data2google.json'), JSON.stringify(dataGoogle, null, '\t'), 'utf8', _ => 1)
    // 4. on synchronise avec Google Calendar
    return calendar.syncCore(dataGoogle)
}
main().then(r => {
        console.log(`\nC'est terminé (${r.length} événements ajoutés/supprimés) ! Béni sois-tu Seigneur !`);
        systemPause()
    })
    .catch(err => {
        console.log('ERREUR', err, '\n\n');
        systemPause()
    })



// 1. renvoie le chemin du fichier le plus récent dans dossier qui matche l'expression régulière regex_filter
function getMostRecentFile(dossier, regex_filter = null) {
    let files = fs.readdirSync(dossier).filter(f => regex_filter.test(f))
    files = mapAddAttr(files, 'lastModified', f => {
        let stats = fs.statSync(path.join(dossier, f));
        return moment(util.inspect(stats.mtime)).format('YYYYMMDDHHmm')
    })
    files = sortBy(files, f => f.lastModified)
    if (files.length < 1) return '';
    return path.join(dossier, files[files.length-1].val)
}

// 2. renvoie les données intéressantes du fichier excel extrait de Logmis
function getData(excel_path) {
    const result = excelToJson({sourceFile: excel_path});
    if (Object.getOwnPropertyNames(result).length < 1) {
        return {error: true, description: `Le fichier excel ${excel_path} semble vide : `}
    }
    // 2.1 on récupère la première worksheet
    const lignes = result[Object.getOwnPropertyNames(result)[0]]
    const data = lignes
                // on filtre les dossiers qui n'ont pas de date de départ ou d'arrivée
                .filter(lig => lig.F && typeof lig.F == 'object' && lig.G && typeof lig.G == 'object')
                // on filtre les dossiers "Contrat annulé"
                .filter(lig => lig.D != 'Contrat annulé')
                .map(lig => {
                    return {
                        dossier: lig.A,
                        client: lig.B,
                        nom_groupe: lig.C,
                        etat: lig.D,
                        date_debut: moment(lig.F.toISOString()),
                        date_fin: moment(lig.G.toISOString()),
                        rooms: lig.H,
                        observation: lig.N,
                        pax: REGEX_PAX(lig),
                        pension: REGEX_PENSION(lig)
                    }
                })
                // on filtre les dossiers antérieurs à aujourd'hui
                .filter(el => el.date_debut.isSameOrAfter(moment().startOf('day')))
    return data
}

// 3. transforme les données issues du fichier excel dans le format google calendar
function data2GoogleCalendarFormat(data) {
    let events = data.map(parseDossier4GoogleCalendar)
    return flatten(events)
}

function parseDossier4GoogleCalendar(dossier) {
    // on ajoute les pti-déj si HB ou BB (pas bivouac)
    let breakfasts = []
    if (dossier.pension && dossier.pension != 'BIVOUAC') {
        breakfasts = getDateRange(dossier.date_debut, dossier.date_fin, HEURE_PTIDEJ).slice(1).map(currTime => {
            return newEvent(dossier, {title: 'Ptidéj -', time: currTime, color: COLORS.breakfast})
        })
    }
    // on ajoute éventuellement les dîners
    let dinners = []
    if (dossier.pension && dossier.pension == 'HB') {
        dinners = getDateRange(dossier.date_debut, dossier.date_fin, HEURE_DINER).slice(0, -1).map(currTime => {
            return newEvent(dossier, {title: 'Dîner -', time: currTime, color: COLORS.dinner})
        })
    }
    // on renvoie tous ces événements + arrivée et départ
    let time_debut = moment(dossier.date_debut)
    time_debut.set({hour: HEURE_ARRIVEE, minutes:0})
    let time_fin = moment(dossier.date_fin)
    time_fin.set({hour: HEURE_DEPART, minutes:0})

    return [newEvent(dossier, {title: 'ARRIVÉE -', date: dossier.date_debut, color: COLORS.arrivee}),
            newEvent(dossier, {title: 'ARRIVÉE -', time: time_debut, color: COLORS.arrivee}), 
            breakfasts, 
            dinners, 
            newEvent(dossier, {title: 'DÉPART -', date: dossier.date_fin, color: COLORS.depart}),
            newEvent(dossier, {title: 'DÉPART -', time: time_fin, color: COLORS.depart})
        ]
}

function newEvent(dossier, opt) {
    let ifpax = (dossier.pax) ? dossier.pax + 'pax ': '';
    let ifoption = (dossier.etat == 'Option') ? 'OPTION -': '';
    let nom_groupe = (dossier.nom_groupe) ? dossier.nom_groupe: dossier.client;
    let ifrooms = (dossier.rooms) ? '\n' + dossier.rooms + ' chambres': '';

    let ev = {
        summary: `${ifoption} ${opt.title} ${nom_groupe} ${ifpax}`.trim(),
        description: `Dossier N°${dossier.dossier} - ${dossier.etat}\n${dossier.client}${ifrooms}`,
        colorId: opt.color,
        start: {},
        end: {}
    }

    if (opt.time) {
        ev.start.dateTime = opt.time.toISOString();
        ev.end.dateTime = opt.time.add(1, 'hours').toISOString()
    } else {
        ev.start.date = opt.date.format('YYYY-MM-DD');
        ev.end.date = opt.date.format('YYYY-MM-DD')
    }
    return ev
}


/* ======================================================= */
//                       HELPERS
/* ======================================================= */

function getDateRange(start_date, end_date, heure = null) {
    let start_date2 = moment(start_date)
    if (heure !== null) start_date2.set({hour:heure, minute:0});
    let res = [moment(start_date2)]
    let nb_days = end_date.diff(start_date, 'days')
    for (let i = 1; i <= nb_days; i++) {
        let currDate = start_date2.add(1, 'days')
        res.push(moment(currDate))
    }
    return res
}

function flatten(arr) {
    let res = []
    for (let el of arr) {
        if (typeof el == 'object' && el.length > 0) {
            res = res.concat(flatten(el))
        } else if (typeof el != 'object' || el.length === undefined) {
            res.push(el)
        }
    }
    return res
}

function mapAddAttr(arr, attr_name, fn) {
    return arr.map(el => {
        let o = el;
        if (typeof el != 'object') o = {val: el};
        o[attr_name] = fn(el)
        return o
    })
}

function sortBy(arr, fn) {
    if (typeof fn == 'string') fn = (el) => el[fn];
    return arr.sort((a,b) => {
      if (fn(a) < fn(b)) return -1;
      return 1
    })
  }

function systemPause() {
    console.log('Appuie sur une touche pour quitter');
    process.stdin.setRawMode(true);
    process.stdin.resume();
    process.stdin.on('data', process.exit.bind(process, 0));
}