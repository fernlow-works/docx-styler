import { beforeEach, describe, expect, it, vi } from "vitest";

interface Call {
	url: string;
	method: string;
	contentType: string;
	form: Record<string, string>;
}
const calls: Call[] = [];
let reply: { status: number; json: unknown } | "throw" = { status: 200, json: {} };

vi.mock("obsidian", () => ({
	requestUrl: (req: { url: string; method: string; contentType: string; body: string }) => {
		calls.push({ url: req.url, method: req.method, contentType: req.contentType, form: Object.fromEntries(new URLSearchParams(req.body)) });
		if (reply === "throw") return Promise.reject(new Error("net::ERR_INTERNET_DISCONNECTED"));
		return Promise.resolve(reply);
	},
}));

import { cacheGrantsPro, normalizeLicense, toCache, verifyWithGumroad } from "../src/license";
import { ACTIVATION_LIMIT, GUMROAD_PRODUCT_ID, GUMROAD_PRODUCT_URL, GUMROAD_VERIFY_URL, PRO_PRICE_LABEL, proIsOnSale } from "../src/product";

const PRODUCT = "test-product-id==";
const KEY = "A1B2C3D4-E5F60718-9ABCDEF0-1234ABCD";

/** Shaped like the documented verify reply, including buyer data that must not be stored. */
function okReply(uses: number, flags: Partial<Record<"refunded" | "disputed" | "chargebacked", boolean>> = {}) {
	return {
		status: 200,
		json: {
			success: true,
			uses,
			purchase: {
				product_id: PRODUCT,
				email: "buyer@example.com",
				license_key: KEY,
				sale_id: "S1",
				refunded: false,
				disputed: false,
				dispute_won: false,
				chargebacked: false,
				...flags,
			},
		},
	};
}

beforeEach(() => {
	calls.length = 0;
	reply = okReply(1);
});

describe("product", () => {
	it("is on sale with the Gumroad product id and page", () => {
		expect(GUMROAD_PRODUCT_ID).toBe("Yuvgn4Oyn8g4Di9P2c2-Yg==");
		expect(GUMROAD_PRODUCT_URL).toBe("https://fernlowworks.gumroad.com/l/docxstylerpro");
		expect(proIsOnSale()).toBe(true);
		expect(PRO_PRICE_LABEL).toBe("19 USD");
		expect(ACTIVATION_LIMIT).toBe(5);
		expect(GUMROAD_VERIFY_URL).toBe("https://api.gumroad.com/v2/licenses/verify");
	});
	it("does not call the network without a product id", async () => {
		expect(await verifyWithGumroad(KEY, true, "")).toBe("not-on-sale");
		expect(await verifyWithGumroad(KEY, false, "")).toBe("not-on-sale");
		expect(calls).toHaveLength(0);
	});
});

describe("verifyWithGumroad request", () => {
	it("Activate posts product_id, license_key and increment_uses_count=true as a form", async () => {
		await verifyWithGumroad(KEY, true, PRODUCT);
		expect(calls).toEqual([
			{
				url: "https://api.gumroad.com/v2/licenses/verify",
				method: "POST",
				contentType: "application/x-www-form-urlencoded",
				form: { product_id: PRODUCT, license_key: KEY, increment_uses_count: "true" },
			},
		]);
	});
	it("uses the shipped Gumroad product id by default", async () => {
		await verifyWithGumroad(KEY, false);
		expect(calls[0]?.form.product_id).toBe("Yuvgn4Oyn8g4Di9P2c2-Yg==");
	});
	it("Check license sends increment_uses_count=false", async () => {
		await verifyWithGumroad(KEY, false, PRODUCT);
		expect(calls[0]?.form).toEqual({ product_id: PRODUCT, license_key: KEY, increment_uses_count: "false" });
	});
});

describe("verifyWithGumroad replies", () => {
	it("activates a good key and stores only status, uses and checkedAt", async () => {
		reply = okReply(1);
		const r = await verifyWithGumroad(KEY, true, PRODUCT);
		expect(r).toMatchObject({ status: "active", uses: 1 });
		expect(Object.keys(r as object).sort()).toEqual(["checkedAt", "status", "uses"]);
		expect(JSON.stringify(r)).not.toContain("buyer@example.com");
		expect(JSON.stringify(r)).not.toContain(KEY);
	});
	it.each(["refunded", "disputed", "chargebacked"] as const)("rejects a %s purchase", async (flag) => {
		reply = okReply(1, { [flag]: true });
		expect(await verifyWithGumroad(KEY, true, PRODUCT)).toMatchObject({ status: flag });
		expect(await verifyWithGumroad(KEY, false, PRODUCT)).toMatchObject({ status: flag });
	});
	it("treats a 404 as an invalid key", async () => {
		reply = { status: 404, json: { success: false, message: "That license does not exist for the provided product." } };
		expect(await verifyWithGumroad(KEY, true, PRODUCT)).toMatchObject({ status: "invalid", uses: null });
	});
	it("treats success false as invalid", async () => {
		reply = { status: 200, json: { success: false } };
		expect(await verifyWithGumroad(KEY, false, PRODUCT)).toMatchObject({ status: "invalid" });
	});
	it("rejects Activate when uses goes over the device limit", async () => {
		reply = okReply(ACTIVATION_LIMIT);
		expect(await verifyWithGumroad(KEY, true, PRODUCT)).toMatchObject({ status: "active", uses: ACTIVATION_LIMIT });
		reply = okReply(ACTIVATION_LIMIT + 1);
		expect(await verifyWithGumroad(KEY, true, PRODUCT)).toMatchObject({ status: "over-limit", uses: ACTIVATION_LIMIT + 1 });
	});
	it("does not apply the limit on Check license, which adds no use", async () => {
		reply = okReply(ACTIVATION_LIMIT + 3);
		expect(await verifyWithGumroad(KEY, false, PRODUCT)).toMatchObject({ status: "active" });
	});
	it("rejects Activate when the reply has no uses count", () => {
		expect(toCache({ success: true, purchase: { refunded: false, disputed: false, chargebacked: false } }, true, 1).status).toBe("over-limit");
	});
	it("treats missing purchase flags as not valid", () => {
		expect(toCache({ success: true, uses: 1, purchase: {} }, false, 1).status).not.toBe("active");
		expect(toCache({ success: true, uses: 1 }, false, 1).status).toBe("invalid");
	});
	it("reports offline for network errors and server errors and stores nothing", async () => {
		reply = "throw";
		expect(await verifyWithGumroad(KEY, true, PRODUCT)).toBe("offline");
		reply = { status: 500, json: null };
		expect(await verifyWithGumroad(KEY, true, PRODUCT)).toBe("offline");
		reply = { status: 400, json: { success: false } };
		expect(await verifyWithGumroad(KEY, false, PRODUCT)).toBe("offline");
	});
});

describe("stored license", () => {
	it("grants Pro offline only for an active result", () => {
		expect(cacheGrantsPro({ status: "active", uses: 1, checkedAt: 1 })).toBe(true);
		expect(cacheGrantsPro({ status: "refunded", uses: 1, checkedAt: 1 })).toBe(false);
		expect(cacheGrantsPro({ status: "over-limit", uses: 6, checkedAt: 1 })).toBe(false);
		expect(cacheGrantsPro(null)).toBe(false);
	});
	it("treats an older license cache shape as no license", () => {
		expect(normalizeLicense({ status: "granted", expiresAt: null, activationId: "act-1", checkedAt: 1 })).toBeNull();
		expect(normalizeLicense({ status: "active" })).toBeNull();
		expect(normalizeLicense("active")).toBeNull();
		expect(normalizeLicense(null)).toBeNull();
		expect(normalizeLicense({ status: "active", uses: 2, checkedAt: 5, email: "x@example.com" })).toEqual({ status: "active", uses: 2, checkedAt: 5 });
	});
});
