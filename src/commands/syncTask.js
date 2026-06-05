import inquirer from "inquirer"
import xlsx from "xlsx"
import attendeeEmails from "../../data/attendeeEmails.json" with { type: "json" }

const games = []
let workbook

export const syncTask = async (options) => {
  // define the usage prompts
  const questions = [
    {
      type: "input", // text input
      name: "filename", // will be saved in answers.filename
      message: "Enter the filename to the scheduling spreadsheet:",
      default: process.env.FILENAME,
      validate: (input) => (input ? true : "Filename cannot be empty"), // input validation
    },
    {
      type: "number", // number input
      name: "year", // will be saved in answers.year
      message: "Enter the schedule year number:",
      default: new Date().getFullYear(),
      validate: (input) => (input ? true : "Year number cannot be empty"), // input validation
    },
    {
      type: "input", // text input
      name: "timeZone", // will be saved in answers.timeZone
      message: "Enter the schedule time zone",
      default: "Australia/Sydney",
      validate: (input) => (input ? true : "Time zone cannot be empty"), // input validation
    },
    {
      type: "input", // text input
      name: "calendarId", // will be saved in answers.calendarId
      message:
        "Enter the Calendar ID - retrieve via Google Calendar API `https://www.googleapis.com/calendar/v3/users/me/calendarList`",
      default: process.env.CALENDAR_ID,
      validate: (input) => (input ? true : "Calendar ID cannot be empty"), // input validation
    },
    {
      type: "input", // text input
      name: "token", // will be saved in answers.token
      message:
        "Enter the Google Calendar API access token - goto `https://developers.google.com/oauthplayground/#step1&scopes=https%3A%2F%2Fwww.googleapis.com%2Fauth%2Fcalendar.events.owned&url=https%3A%2F%2F&content_type=application%2Fjson&http_method=GET&useDefaultOauthCred=unchecked&oauthEndpointSelect=Google&oauthAuthEndpointValue=https%3A%2F%2Faccounts.google.com%2Fo%2Foauth2%2Fv2%2Fauth&oauthTokenEndpointValue=https%3A%2F%2Foauth2.googleapis.com%2Ftoken&includeCredentials=unchecked&accessTokenType=bearer&autoRefreshToken=unchecked&accessType=offline&prompt=consent&response_type=code&wrapLines=on`",
      default: process.env.ACCESS_TOKEN,
      validate: (input) => (input ? true : "Access token cannot be empty"), // input validation
    },
  ]

  try {
    // show questions to the user and wait for answers
    const answers = await inquirer.prompt(questions)
    // print answers as confirmation
    console.log("\n🏒 schedule sync task added OK!")
    console.log("\n🏒 sync task details:", answers)
    // read the spreadsheet using supplied filename
    readSpreadsheet(answers, options)
    // sync games to the google calendar
    syncGamesToCalendar(answers, options)
  } catch (error) {
    console.error("🏒 syncTask error:", error)
  }
}

const readSpreadsheet = (answers, options) => {
  try {
    workbook = xlsx.readFile(answers.filename, {
      cellDates: true,
    })
    const sheet = workbook.Sheets[workbook.SheetNames[0]]
    let nextRow = 2
    let foundNextGame = sheet[`C${nextRow}`]?.v ? true : false
    while (foundNextGame) {
      const game = {
        date: sheet[`C${nextRow}`]?.v,
        start: sheet[`E${nextRow}`]?.w,
        finish: sheet[`F${nextRow}`]?.w,
        league: sheet[`D${nextRow}`]?.w,
        home: sheet[`G${nextRow}`]?.w,
        away: sheet[`H${nextRow}`]?.w,
        ref1: sheet[`I${nextRow}`]?.w,
        ref2: sheet[`J${nextRow}`]?.w,
        liner1: sheet[`K${nextRow}`]?.w,
        liner2: sheet[`L${nextRow}`]?.w,
        scorer1: sheet[`M${nextRow}`]?.w,
        scorer2: sheet[`N${nextRow}`]?.w,
      }
      const month = new Date(Date.parse(game.date)).getMonth() + 1
      const paddedMonth = month.toString().padStart(2, "0")
      const paddedDay = new Date(Date.parse(game.date))
        .getDate()
        .toString()
        .padStart(2, "0")
      const calendarStartDateTime = `${answers.year}-${paddedMonth}-${paddedDay}T${game.start}:00`
      game.calendarStart = {
        dateTime: calendarStartDateTime,
        timeZone: answers.timeZone,
      }
      const calendarEndDateTime = `${answers.year}-${paddedMonth}-${paddedDay}T${game.finish}:00`
      game.calendarEnd = {
        dateTime: calendarEndDateTime,
        timeZone: answers.timeZone,
      }
      console.debug("🏒 game:", game)
      games.push(game)
      nextRow++
      foundNextGame = sheet[`C${nextRow}`]?.v ? true : false
    }
  } catch (error) {
    console.error("🏒 readSpreadsheet error:", error)
  }
}

const syncGamesToCalendar = async (answers, options) => {
  try {
    const headers = new Headers()
    headers.set("Accept", "application/json")
    headers.set("Authorization", `Bearer ${answers.token}`)
    const eventsRequest = new Request(
      `https://www.googleapis.com/calendar/v3/calendars/${answers.calendarId}/events`,
      {
        headers: headers,
      },
    )
    const eventsResponse = await fetch(eventsRequest)
    if (!eventsResponse.ok)
      throw new Error(
        `HTTP error! status: ${eventsResponse.status} ${eventsResponse.statusText}`,
      )
    const data = await eventsResponse.json()
    // console.debug("🏒 syncGamesToCalendar - events list response:", data)
    console.log("Comparing games to events for calendar: ", data.summary)
    const events = data.items || []
    for (const game of games) {
      const eventSummary = `🏒 Officiating - ${game.league}: ${game.away || "[Unknown]"} @ ${game.home || "[Unknown]"}`
      const eventDescription = `Referee 1: ${game.ref1 || "[None]"}
Referee 2: ${game.ref2 || "[None]"}
Linesperson 1: ${game.liner1 || "[None]"}
Linesperson 2: ${game.liner2 || "[None]"}
Scorer 1: ${game.scorer1 || "[None]"}
Scorer 2: ${game.scorer2 || "[None]"}
`
      // check if any of the officials are configured with an email address
      // if so, add them as attendees so they receive an invite
      const attendees = []
      for (const attendeeEmail of attendeeEmails) {
        if (
          (game.ref1 &&
            game.ref1.trim().toLowerCase() ===
              attendeeEmail.displayName.toLowerCase()) ||
          (game.ref2 &&
            game.ref2.trim().toLowerCase() ===
              attendeeEmail.displayName.toLowerCase()) ||
          (game.liner1 &&
            game.liner1.trim().toLowerCase() ===
              attendeeEmail.displayName.toLowerCase()) ||
          (game.liner2 &&
            game.liner2.trim().toLowerCase() ===
              attendeeEmail.displayName.toLowerCase()) ||
          (game.scorer1 &&
            game.scorer1.trim().toLowerCase() ===
              attendeeEmail.displayName.toLowerCase()) ||
          (game.scorer2 &&
            game.scorer2.trim().toLowerCase() ===
              attendeeEmail.displayName.toLowerCase())
        ) {
          attendees.push(attendeeEmail)
        }
      }
      // find existing event for game start and end times
      const event = events.find(
        (event) =>
          event.start.dateTime.indexOf(game.calendarStart.dateTime) > -1,
      )
      if (event) {
        // update existing event
        let eventHasChanges = false
        if (
          event.summary !== eventSummary ||
          event.description !== eventDescription
        ) {
          event.summary = eventSummary
          event.description = eventDescription
          eventHasChanges = true
        }
        if (attendees.length > 0) {
          // compare attendees
          if (!event.attendees) {
            eventHasChanges = true
          } else {
            for (const attendee of attendees) {
              const eventAttendee = event.attendees?.find(
                (eventAttendee) => eventAttendee.email === attendee.email,
              )
              if (!eventAttendee) {
                // found new attendee not yet in event
                eventHasChanges = true
              } else {
                attendee.responseStatus = eventAttendee.responseStatus
              }
            }
            for (const eventAttendee of event.attendees) {
              const attendee = attendees.find(
                (attendee) => attendee.email === eventAttendee.email,
              )
              if (!attendee) {
                // found event attendee to be removed
                eventHasChanges = true
              }
            }
          }
          if (eventHasChanges) {
            event.attendees = attendees.slice()
          }
        } else {
          if (event.attendees && event.attendees.length > 0) {
            // no attendees - remove existing attendees
            event.attendees = []
            eventHasChanges = true
          }
        }
        if (eventHasChanges) {
          if (!options.test) {
            const eventUpdateRequest = new Request(
              `https://www.googleapis.com//calendar/v3/calendars/${answers.calendarId}/events/${event.id}?sendUpdates=all`,
              {
                headers: headers,
                method: "PUT",
                body: JSON.stringify(event),
              },
            )
            const eventUpdateResponse = await fetch(eventUpdateRequest)
            if (!eventUpdateResponse.ok)
              throw new Error(
                `HTTP error! status: ${eventUpdateResponse.status} ${eventUpdateResponse.statusText}`,
              )
            const updatedEvent = await eventUpdateResponse.json()
            console.debug(
              "🏒 syncGamesToCalendar - updated event response:",
              updatedEvent,
            )
          } else {
            console.debug(
              "🏒 syncGamesToCalendar - event to update - test only:",
              event,
            )
          }
        } else {
          console.debug(
            "🏒 syncGamesToCalendar - event unchanged",
            event.summary,
            event.start.dateTime,
            event.end.dateTime,
          )
        }
      } else {
        // insert new event
        const newEvent = {
          summary: eventSummary,
          description: eventDescription,
          start: {
            dateTime: game.calendarStart.dateTime,
            timeZone: game.calendarStart.timeZone,
          },
          end: {
            dateTime: game.calendarEnd.dateTime,
            timeZone: game.calendarEnd.timeZone,
          },
        }
        if (attendees.length > 0) {
          newEvent.attendees = attendees.slice()
        }
        if (!options.test) {
          const eventInsertRequest = new Request(
            `https://www.googleapis.com//calendar/v3/calendars/${answers.calendarId}/events?sendUpdates=all`,
            {
              headers: headers,
              method: "POST",
              body: JSON.stringify(newEvent),
            },
          )
          const eventInsertResponse = await fetch(eventInsertRequest)
          if (!eventInsertResponse.ok)
            throw new Error(
              `HTTP error! status: ${eventInsertResponse.status} ${eventInsertResponse.statusText}`,
            )
          const insertedEvent = await eventInsertResponse.json()
          console.debug(
            "🏒 syncGamesToCalendar - inserted event response:",
            insertedEvent,
          )
        } else {
          console.debug(
            "🏒 syncGamesToCalendar - event to insert - test only:",
            newEvent,
          )
        }
      }
    }
  } catch (error) {
    console.error("🏒 syncGamesToCalendar error:", error)
  }
}
