// @vitest-environment happy-dom
import { describe, expect, it } from "vitest"
import { createIcon, iconButton } from "./chat-icons"

describe("context action icons", () => {
  it.each<"help" | "settings">(["help", "settings"])(
    "publishes stable %s icon metadata",
    (name) => {
      const icon = createIcon(name)
      expect(icon.getAttribute("data-icon")).toBe(name)
      expect(icon.getAttribute("aria-hidden")).toBe("true")
    },
  )

  it("draws settings as a toothed gear with a center hole", () => {
    const icon = createIcon("settings")
    expect(icon.querySelector('[data-icon-part="gear-teeth"]')).not.toBeNull()
    expect(icon.querySelector('[data-icon-part="gear-hole"]')).not.toBeNull()
  })

  it("gives an icon button an accessible name and title", () => {
    const button = iconButton("설정", "settings", () => {})
    expect(button.getAttribute("aria-label")).toBeTruthy()
    expect(button.getAttribute("title")).toBeTruthy()
  })
})
