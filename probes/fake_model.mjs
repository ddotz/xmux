// A stand-in for the KDB AI server (OpenAI **chat completions**), used to exercise the 대화 tab end to end
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

/**
 * The real server answers 405 on anything but this route. An earlier version of this fake
 * accepted every path and answered in the legacy `completions` shape, so the pane's wrong
 * route reached production unnoticed. It is strict on purpose now.
 */
const ROUTE = "/api/chat/completions"

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

  const path = new URL(request.url ?? "/", "https://localhost").pathname
  if (path !== ROUTE || request.method !== "POST") {
    console.log(`${request.method} ${path} -> 405 (route is POST ${ROUTE})`)
    response.writeHead(405, { ...cors, "Content-Type": "application/json" })
    response.end(JSON.stringify({ detail: "Method Not Allowed" }))
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

    // The real server refuses a system message anywhere but first, and refuses more than
    // one. Answering 200 regardless is what let that reach production once already.
    let messages = []
    try {
      messages = JSON.parse(body).messages ?? []
    } catch {
      messages = []
    }
    const misplaced = messages.some((message, index) => message.role === "system" && index !== 0)
    if (misplaced) {
      console.log("-> 400 (system message must be at the beginning)")
      response.writeHead(400, { ...cors, "Content-Type": "application/json" })
      response.end(
        JSON.stringify({ detail: "400: System message must be at the beginning." }),
      )
      return
    }

    response.writeHead(200, { ...cors, "Content-Type": "application/json" })
    response.end(
      JSON.stringify({
        choices: [{ message: { role: "assistant", content: ANSWER } }],
        usage: { completion_tokens: 42 },
      }),
    )
  })
}).listen(port, () => {
  console.log(`fake model on https://localhost:${port}`)
})
