/**
 * Typed FlipBox class builder — recreates the validation half of the web repo's React
 * FlipBox component with zero runtime: the unions make invalid layouts unrepresentable.
 * The CSS itself (flipbox.css) is vendored verbatim from ~/Developer/web.
 */

export type FbDirection = 'x' | 'y'
export type FbAlign = 'N' | 'S' | 'E' | 'W' | 'NE' | 'NW' | 'SE' | 'SW' | 'NS' | 'WE' | 'C'
export type FbSizing = 'x-hug' | 'x-fill' | 'x-fixed' | 'y-hug' | 'y-fill' | 'y-fixed'
export type FbExtra = 'wrap' | 'nowrap' | 'abs' | 'paragraph' | 'text' | `gap-${1 | 2 | 3}` | `pad-${2 | 3}`

export const fb = (direction: FbDirection, align: FbAlign, ...rest: Array<FbSizing | FbExtra>): string =>
  ['flipbox', direction, align, ...rest].join(' ')
