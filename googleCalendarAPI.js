const ARCANA = require('./arcana.json')
const KEYS = ARCANA.service_account
const CALENDAR_ID = ARCANA.CALENDAR_ID

const {JWT} = require('google-auth-library')
const moment = require('moment')
let GCLIENT;

// COLOR IDS : https://eduardopereira.pt/wp-content/uploads/2012/06/google_calendar_api_event_color_chart.png

// synchronise les événements de list_events dans le calendrier
// on peut l'améliorer, sans utiliser la fontion segment() mais en faisant un await dès que compteur = nb_batch...
async function syncCoreBatch(list_events, nb_batch = 30) {
  let existing_events = await getFutureEvents()
  let all_res = [];

  // on ajoute les événements modifiés ou ajoutés
  let batches_list_events = segment(list_events, nb_batch);

  for (let i = 0; i < batches_list_events.length; i++) {
    console.log(`\nBatch ADD ${i+1}/${batches_list_events.length}`)
    let myPromises = []
    let compteur = 0
    let status = 0

    for (let ev of batches_list_events[i]) {
      if (!eventExistsInList(ev, existing_events)) {
        let p = addEvent(ev, 10)
        p.then(_ => {
          status++;
          printText(`Sync Status : ${status}/${compteur}`)
        })
        p.catch(e => console.log(e))
        compteur++;
        printText(compteur.toString())
        myPromises.push(p)
      }
    }

    let curr_res = await Promise.all(myPromises)
    all_res = all_res.concat(curr_res)
  }

  // on supprime les événements en trop
  let batches_del = segment(existing_events, nb_batch);
  for (let i = 0; i < batches_del.length; i++) {
    console.log(`\nBatch DEL ${i+1}/${batches_del.length}`)
    let myPromises = []
    let compteur = 0
    let status = 0

    for (let ev of existing_events) {
      if (!eventExistsInList(ev, list_events)) {
        let p = deleteEvent(ev, 10)
        p.then(_ => {
          status++;
          printText(`Sync Status : ${status}/${compteur}`)
        })
        p.catch(e => console.log(e))
        compteur++;
        printText(compteur.toString())
        myPromises.push(p)
      }
    }

    let curr_res = await Promise.all(myPromises)
    all_res = all_res.concat(curr_res)
  }

  return all_res
}

// synchronise les événements de list_events dans le calendrier
async function syncCore(list_events) {
  let existing_events = await getFutureEvents()
  let myPromises = []
  let compteur = 0
  let status = 0
  // on ajoute les événements modifiés ou ajoutés
  list_events.forEach(ev => {
    if (!eventExistsInList(ev, existing_events)) {
      let p = addEvent(ev, 10)
      p.then(_ => {
        status++;
        printText(`Sync Status : ${status}/${compteur}`)
      })
      p.catch(e => console.log(e))
      compteur++;
      printText(compteur.toString())
      myPromises.push(p)
    }
  })
  // on supprime les événements en trop
  existing_events.forEach(ev => {
    if (!eventExistsInList(ev, list_events)) {
      let p = deleteEvent(ev, 10)
      p.then(_ => {
        status++;
        printText(`Sync Status : ${status}/${compteur}`)
      })
      p.catch(e => console.log(e))
      compteur++;
      printText(compteur.toString())
      myPromises.push(p)
    }
  })

  return Promise.all(myPromises)
}

async function getFutureEvents() {
    GCLIENT = await authenticate()
  
    let res = await GCLIENT.request({
      url: `https://www.googleapis.com/calendar/v3/calendars/${CALENDAR_ID}/events`,
      params: {
        calendarId: CALENDAR_ID,
        timeMin: moment().startOf('day').toISOString(),
        maxResults: 1000,
        singleEvents: true,
        orderBy: 'startTime'
      }
    })
    return res.data.items
}
  
async function addEvent(ev, retry = 0) {
    if (retry < 0) throw 'Impossible to add event (number of retries exceeded) : ' + JSON.stringify(ev)
    GCLIENT = await authenticate()
    await delay(Math.random() * 1000)
    try {
      let res = await GCLIENT.request({
        url: `https://www.googleapis.com/calendar/v3/calendars/${CALENDAR_ID}/events`,
        method: 'POST',
        data: ev
      })
      return res.data
    } catch (e) {
      if (retry > 0) return addEvent(ev, retry - 1)
      else throw e
    }
}
  
async function deleteEvent(ev, retry = 0) {
    if (retry < 0) throw 'Impossible to delete event (number of retries exceeded) : ' + JSON.stringify(ev)
    GCLIENT = await authenticate()
    let eventId = (ev.id || ev)
    if (!eventId) throw "Invalid event - Unable to delete : " & JSON.stringify(ev)
    await delay(Math.random() * 1000)
    try {
      let res = await GCLIENT.request({
        url: `https://www.googleapis.com/calendar/v3/calendars/${CALENDAR_ID}/events/${eventId}`,
        method: 'delete'
      })
      return res
    } catch (e) {
      if (retry > 0) return deleteEvent(ev, retry - 1)
      else throw e
    }
  }
  
async function authenticate() {
    if (GCLIENT) return GCLIENT;
    GCLIENT = new JWT(
      KEYS.client_email,
      null,
      KEYS.private_key, ['https://www.googleapis.com/auth/calendar'],
    );
    await GCLIENT.authorize();
    return GCLIENT
}
  
  // dit si l'événement ev existe dans la liste d'événements list_events
function eventExistsInList(ev, list_events) {
    for (e of list_events) {
      if (eventsEqual(e, ev)) return true
    }
    return false
}
  
function eventsEqual(ev1, ev2) {
  let b = ev1.summary == ev2.summary && ev1.description == ev2.description &&
    ev1.start && ev2.start && ev1.end && ev2.end;
  if (ev1.start.dateTime) {
    return b && ev1.start.dateTime && ev1.end.dateTime && ev2.start.dateTime && ev2.end.dateTime &&
      moment(ev1.start.dateTime).isSame(ev2.start.dateTime) &&
      moment(ev1.end.dateTime).isSame(ev2.end.dateTime)
  } else {
    return b && ev1.start.date && ev1.end.date && ev2.start.date && ev2.end.date &&
      ev1.start.date == ev2.start.date && ev1.end.date == ev2.end.date
  }
}

function delay(ms) {
  return new Promise((resolve, reject) => {
    setTimeout(resolve, ms)
  })
}

function printText(texte) {
  process.stdout.clearLine();
  process.stdout.cursorTo(0);
  process.stdout.write(texte);
}

// renvoie un array 2d avec les éléments de arr répartis par batches de @nb éléments
function segment(arr, nb) {
  let res = []
  let batch = []
  for (let i = 0; i < arr.length; i++) {
    if (i % nb == 0) {
      if (batch.length > 0) res.push(batch)
      batch = [];
    }
    batch.push(arr[i])
  }
  if (batch.length > 0) res.push(batch)
  return res
}

module.exports = {
  deleteEvent,
  addEvent,
  getFutureEvents,
  syncCore,
  syncCoreBatch
}