import type { HiddenRange } from './types.js'

export interface NormalizedRanges {
	// Sorted, deduplicated, 1-based indices (Excel's numbering).
	indices: number[]
	// Coalesced [min, max] intervals over the sorted indices (1-based, inclusive).
	intervals: Array<[number, number]>
}

export default function normalizeRanges(
	ranges: HiddenRange[] | undefined,
	kind: 'row' | 'column',
	max: number
): NormalizedRanges {
	const set = new Set<number>()

	if (!ranges) {
		return { indices: [], intervals: [] }
	}

	for (const range of ranges) {
		if (typeof range === 'number') {
			assertValidIndex(range, kind, max)
			set.add(range + 1)
		} else if (range && typeof range === 'object' && 'from' in range && 'to' in range) {
			const { from, to } = range
			assertValidIndex(from, kind, max)
			assertValidIndex(to, kind, max)
			if (from > to) {
				throw new Error(
					`hidden${kind === 'row' ? 'Rows' : 'Columns'} range has from > to: { from: ${from}, to: ${to} }`
				)
			}
			for (let i = from; i <= to; i++) {
				set.add(i + 1)
			}
		} else {
			throw new Error(
				`hidden${kind === 'row' ? 'Rows' : 'Columns'} entries must be a number or { from, to } object`
			)
		}
	}

	const indices = Array.from(set).sort((a, b) => a - b)
	const intervals: Array<[number, number]> = []
	for (const n of indices) {
		const last = intervals[intervals.length - 1]
		if (last && n === last[1] + 1) {
			last[1] = n
		} else {
			intervals.push([n, n])
		}
	}
	return { indices, intervals }
}

function assertValidIndex(value: number, kind: 'row' | 'column', max: number): void {
	if (!Number.isInteger(value)) {
		throw new Error(
			`hidden${kind === 'row' ? 'Rows' : 'Columns'} indices must be integers (got ${value})`
		)
	}
	if (value < 0) {
		throw new Error(
			`hidden${kind === 'row' ? 'Rows' : 'Columns'} indices must be >= 0 (got ${value})`
		)
	}
	if (value >= max) {
		throw new Error(
			`hidden${kind === 'row' ? 'Rows' : 'Columns'} index ${value} exceeds Excel's maximum ${kind} index (${max - 1})`
		)
	}
}
