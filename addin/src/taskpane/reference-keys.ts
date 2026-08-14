export type ReferenceKeyAction =
  | { readonly kind: "cycle"; readonly step: -1 | 1 }
  | { readonly kind: "jump" }
  | { readonly kind: "back" }
  | { readonly kind: "delete" }

export const keyToReferenceAction = (key: string): ReferenceKeyAction | null => {
  switch (key) {
    case "ArrowLeft":
      return { kind: "cycle", step: -1 }
    case "ArrowRight":
      return { kind: "cycle", step: 1 }
    case "Enter":
      return { kind: "jump" }
    case "Escape":
      return { kind: "back" }
    case "Delete":
    case "Backspace":
      return { kind: "delete" }
    default:
      return null
  }
}
