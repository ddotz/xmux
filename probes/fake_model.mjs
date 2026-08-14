// A stand-in for the KDB AI server (OpenAI **completions**: one prompt in, one text out), used to exercise the 대화 tab end to end
// without spending a real API key. It speaks exactly the one route the pane calls and
// answers with a canned plan, so what is under test is the pane, the request, the
// approval step and the Excel write — everything except the model itself.
//
// Run: node probes/fake_model.mjs [port]
// It serves https because the pane is https; mixed content would never leave the browser.

import { createServer } from "node:https"
import { readFileSync } from "node:fs"
import { homedir } from "node:os"
import { join } from "node:path"

const port = Number(process.argv[2] ?? 3100)
const certs = join(homedir(), ".office-addin-dev-certs")
const options = {
  key: readFileSync(join(certs, "localhost.key")),
  cert: readFileSync(join(certs, "localhost.crt")),
}

/** What the fake model always answers: prose plus one JSON block of edits. */
const ANSWER = [
  "Main!B9에 Data 시트의 합계를 넣겠습니다.",
  "```json",
  '{"edits":[{"sheet":"Main","address":"B9","value":"=SUM(Data!B2:D5)"}]}',
  "```",
].join("\n")

createServer(options, (request, response) => {
  const cors = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "authorization, content-type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
  }
  if (request.method === "OPTIONS") {
    response.writeHead(204, cors)
    response.end()
    return
  }

  let body = ""
  request.on("data", (chunk) => {
    body += chunk
  })
  request.on("end", () => {
    // Printed so the test can confirm the pane really sent the key and the workbook context.
    console.log(`${request.method} ${request.url} auth=${request.headers.authorization ?? "none"}`)
    console.log(`body=${body.slice(0, 400)}`)
    response.writeHead(200, { ...cors, "Content-Type": "application/json" })
    response.end(JSON.stringify({ choices: [{ text: ANSWER }], usage: { completion_tokens: 42 } }))
  })
}).listen(port, () => {
  console.log(`fake model on https://localhost:${port}`)
})
