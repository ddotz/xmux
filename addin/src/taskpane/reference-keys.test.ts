import { describe, expect, it } from "vitest"
import { keyToReferenceAction } from "./reference-keys"

describe("reference keyboard actions", () => {
  it.each([
    ["ArrowLeft", { kind: "cycle", step: -1 }],
    ["ArrowRight", { kind: "cycle", step: 1 }],
    ["Enter", { kind: "jump" }],
    ["Escape", { kind: "back" }],
    ["Delete", { kind: "delete" }],
    ["Backspace", { kind: "delete" }],
  ])("maps %s", (key, expected) => {
    expect(keyToReferenceAction(key)).toEqual(expected)
  })

  it.each(["ArrowUp", "ArrowDown", "Tab", "a"])("does not map %s", (key) => {
    expect(keyToReferenceAction(key)).toBeNull()
  })
})
