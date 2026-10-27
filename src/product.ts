// Everything about the paid tier lives here so the price, product page and
// Gumroad product id are changed in one place. None of these values are
// secrets: Gumroad's license verify endpoint needs no OAuth app or token.

/**
 * Public Gumroad product id, shown in the License key block of the product's
 * content page and in the public product JSON.
 */
export const GUMROAD_PRODUCT_ID = "Yuvgn4Oyn8g4Di9P2c2-Yg==";

/** Public Gumroad product page for Docx Styler Pro. */
export const GUMROAD_PRODUCT_URL = "https://fernlowworks.gumroad.com/l/docxstylerpro";

/** Gumroad license verify endpoint. https://gumroad.com/api (Licenses) */
export const GUMROAD_VERIFY_URL = "https://api.gumroad.com/v2/licenses/verify";

/** Price shown in settings and README. */
export const PRO_PRICE_LABEL = "19 USD";

/** Devices one key may be activated on. Checked against Gumroad's `uses` count on Activate. */
export const ACTIVATION_LIMIT = 5;

export function proIsOnSale(): boolean {
	return GUMROAD_PRODUCT_ID.length > 0 && GUMROAD_PRODUCT_URL.startsWith("https://");
}
