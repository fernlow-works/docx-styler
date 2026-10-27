export type ImageKind = "png" | "jpg" | "gif" | "bmp";

export interface ImageInfo {
	kind: ImageKind;
	width: number;
	height: number;
}

/** Reads width and height from the file header so no DOM or canvas is needed. */
export function readImageInfo(bytes: Uint8Array): ImageInfo | null {
	const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	if (bytes.length > 24 && dv.getUint32(0) === 0x89504e47) {
		return { kind: "png", width: dv.getUint32(16), height: dv.getUint32(20) };
	}
	if (bytes.length > 10 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
		return { kind: "gif", width: dv.getUint16(6, true), height: dv.getUint16(8, true) };
	}
	if (bytes.length > 26 && bytes[0] === 0x42 && bytes[1] === 0x4d) {
		return { kind: "bmp", width: Math.abs(dv.getInt32(18, true)), height: Math.abs(dv.getInt32(22, true)) };
	}
	if (bytes.length > 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
		let i = 2;
		while (i + 9 < bytes.length) {
			if (bytes[i] !== 0xff) {
				i++;
				continue;
			}
			const marker = bytes[i + 1] ?? 0;
			const len = dv.getUint16(i + 2);
			const isSof = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
			if (isSof) return { kind: "jpg", width: dv.getUint16(i + 7), height: dv.getUint16(i + 5) };
			i += 2 + len;
		}
	}
	return null;
}
