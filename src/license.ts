import { requestUrl } from "obsidian";
import { ACTIVATION_LIMIT, GUMROAD_PRODUCT_ID, GUMROAD_VERIFY_URL } from "./product";

/**
 * What we keep after Gumroad answers. Gumroad's reply also contains the
 * buyer's email and other purchase data. None of that is stored.
 */
export interface LicenseCache {
	/**
	 * "active": the key is valid for Pro on this device.
	 * "refunded", "disputed", "chargebacked": Gumroad reports the purchase as such.
	 * "invalid": Gumroad does not know the key for this product.
	 * "over-limit": the key is already in use on more than ACTIVATION_LIMIT devices.
	 */
	status: LicenseStatus;
	/** Gumroad's `uses` count from the last reply, or null if it was missing. */
	uses: number | null;
	/** Epoch ms of the last answer from Gumroad. */
	checkedAt: number;
}

export type LicenseStatus = "active" | "refunded" | "disputed" | "chargebacked" | "invalid" | "over-limit";

const STATUSES: readonly LicenseStatus[] = ["active", "refunded", "disputed", "chargebacked", "invalid", "over-limit"];

export type LicenseResult = LicenseCache | "offline" | "not-on-sale";

export function describeLicense(c: LicenseCache | null): string {
	if (!c) return "No key activated.";
	const when = new Date(c.checkedAt).toLocaleDateString();
	switch (c.status) {
		case "active":
			return `Pro is active. Last checked ${when}.`;
		case "over-limit":
			return `This key is already active on ${ACTIVATION_LIMIT} devices.`;
		case "refunded":
			return "This purchase was refunded, so the key is not active.";
		case "disputed":
		case "chargebacked":
			return "This purchase is disputed, so the key is not active.";
		default:
			return "This key is not active.";
	}
}

/** Offline answer from the stored result. Never touches the network. */
export function cacheGrantsPro(c: LicenseCache | null): boolean {
	return c?.status === "active";
}

/**
 * Settings saved by older versions may hold a license cache of another shape.
 * Anything that is not exactly our shape counts as no license.
 */
export function normalizeLicense(raw: unknown): LicenseCache | null {
	if (!raw || typeof raw !== "object") return null;
	const r = raw as Record<string, unknown>;
	if (typeof r.status !== "string" || !STATUSES.includes(r.status as LicenseStatus)) return null;
	if (typeof r.checkedAt !== "number" || !Number.isFinite(r.checkedAt)) return null;
	const uses = typeof r.uses === "number" && Number.isFinite(r.uses) ? r.uses : null;
	return { status: r.status as LicenseStatus, uses, checkedAt: r.checkedAt };
}

interface GumroadPurchase {
	refunded?: unknown;
	disputed?: unknown;
	chargebacked?: unknown;
}

interface GumroadReply {
	success?: unknown;
	uses?: unknown;
	purchase?: GumroadPurchase;
}

/**
 * Turns a Gumroad 200 reply into what we store. Only `success`, `uses` and the
 * refunded, disputed and chargebacked flags are read. Exported for tests.
 */
export function toCache(reply: GumroadReply, incrementUses: boolean, now = Date.now()): LicenseCache {
	const uses = typeof reply.uses === "number" && Number.isFinite(reply.uses) ? reply.uses : null;
	const p = reply.purchase;
	let status: LicenseStatus;
	if (reply.success !== true || !p) status = "invalid";
	else if (p.refunded !== false) status = "refunded";
	else if (p.chargebacked !== false) status = "chargebacked";
	else if (p.disputed !== false) status = "disputed";
	// Activate counts this device. More uses than the limit means too many devices.
	else if (incrementUses && (uses === null || uses > ACTIVATION_LIMIT)) status = "over-limit";
	else status = "active";
	return { status, uses, checkedAt: now };
}

/**
 * The only network call, made only when the user clicks "Activate"
 * (incrementUses true) or "Check license" (incrementUses false).
 * Sends the public product id, the key and increment_uses_count. Gumroad
 * documents this endpoint as needing no OAuth app or secret.
 * https://gumroad.com/help/article/76-license-keys
 */
export async function verifyWithGumroad(key: string, incrementUses: boolean, productId = GUMROAD_PRODUCT_ID): Promise<LicenseResult> {
	if (!productId) return "not-on-sale";
	try {
		const res = await requestUrl({
			url: GUMROAD_VERIFY_URL,
			method: "POST",
			contentType: "application/x-www-form-urlencoded",
			body: new URLSearchParams({
				product_id: productId,
				license_key: key,
				increment_uses_count: incrementUses ? "true" : "false",
			}).toString(),
			throw: false,
		});
		// Gumroad answers 404 when the key does not verify for this product.
		if (res.status === 404) return { status: "invalid", uses: null, checkedAt: Date.now() };
		if (res.status !== 200) return "offline";
		return toCache(res.json as GumroadReply, incrementUses);
	} catch {
		return "offline";
	}
}
