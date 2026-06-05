#!/usr/bin/env node   // allows this script to be run directly from the terminal

import { Command } from "commander"
import { syncTask } from "./commands/syncTask.js"

const program = new Command() // create new CLI program

// set the CLI app name, description and version
program
  .name("ihact-officiating-calendar-cli")
  .description(
    "Command-line utility to sync the 🏒 officiating schedule spreadsheet to a Google Calendar",
  )
  .version("1.0.0")

// define the 'sync' command that will trigger the interactive prompt
program
  .command("sync")
  .description(
    "Sync the 🏒 officiating schedule spreadsheet to a Google Calendar",
  )
  .option("-t, --test", "Sync in test mode only")
  .action((options) => {
    syncTask(options) // call the sync task function
  })

// parse the command-line arguments
program.parse(process.argv)
