"""Customer-facing shipment wording shared by legacy and fulfillment flows."""

from __future__ import annotations

import html


CUSTOMER_NOTE_TEXT = {
    "PL": {
        "shipped": "Przesyłka dla Twojego zamówienia została nadana za pośrednictwem {carrier}. Numer przesyłki: {tracking}.",
        "cod": " Kwota pobrania dla tej przesyłki: {amount} {currency}.",
        "shipping_included": " Kwota obejmuje pełny koszt dostawy zamówienia w wysokości {shipping} {currency}; pozostałe przesyłki nie będą ponownie obciążone kosztem dostawy.",
        "shipping_not_repeated": " Ta przesyłka obejmuje wyłącznie wartość znajdujących się w niej produktów; koszt dostawy nie zostanie naliczony ponownie.",
        "other_amounts": " W pozostałych przesyłkach płatna będzie wyłącznie wartość zawartych w nich produktów.",
        "no_cod": " Przy odbiorze tej przesyłki nie jest wymagana płatność.",
        "remaining": " Pozostałe produkty z zamówienia zostaną wysłane w osobnej przesyłce.",
    },
    "HU": {
        "shipped": "A rendeléséhez tartozó csomagot átadtuk a(z) {carrier} futárszolgálatnak. Csomagszám: {tracking}.",
        "cod": " A csomag utánvétes összege: {amount} {currency}.",
        "shipping_included": " Ez az összeg tartalmazza a rendelés teljes, {shipping} {currency} összegű szállítási díját; a többi csomagnál nem számítunk fel újabb szállítási díjat.",
        "shipping_not_repeated": " Ez a csomag csak a benne lévő termékek értékét tartalmazza; a szállítási díjat nem számítjuk fel újra.",
        "other_amounts": " A rendelés további csomagjainál csak az azokban található termékek összege fizetendő.",
        "no_cod": " A csomag átvételekor nem kell fizetni.",
        "remaining": " A rendelés többi termékét külön csomagban küldjük.",
    },
    "CZ": {
        "shipped": "Zásilka k vaší objednávce byla odeslána prostřednictvím dopravce {carrier}. Číslo zásilky: {tracking}.",
        "cod": " Dobírka za tuto zásilku činí {amount} {currency}.",
        "shipping_included": " Tato částka zahrnuje celé poštovné objednávky ve výši {shipping} {currency}; u dalších zásilek nebude poštovné účtováno znovu.",
        "shipping_not_repeated": " Tato zásilka zahrnuje pouze hodnotu obsaženého zboží; poštovné nebude účtováno znovu.",
        "other_amounts": " U dalších zásilek bude účtována pouze hodnota zboží, které obsahují.",
        "no_cod": " Při převzetí této zásilky není vyžadována žádná platba.",
        "remaining": " Zbývající produkty z objednávky budou odeslány v samostatné zásilce.",
    },
    "EN": {
        "shipped": "A parcel for your order has been shipped via {carrier}. Tracking number: {tracking}.",
        "cod": " The cash-on-delivery amount for this parcel is {amount} {currency}.",
        "shipping_included": " This amount includes the order's full shipping charge of {shipping} {currency}; no shipping charge will be collected again for the other parcels.",
        "shipping_not_repeated": " This parcel includes only the value of the products it contains; the shipping charge will not be collected again.",
        "other_amounts": " For the other parcels, only the value of the products they contain will be collected.",
        "no_cod": " No payment is required when this parcel is delivered.",
        "remaining": " The remaining products in your order will be shipped in a separate parcel.",
    },
}


def customer_language(country: str | None) -> str:
    value = str(country or "").strip().upper()
    return value if value in {"PL", "HU", "CZ"} else "EN"


def basic_shipment_note(
    country: str | None,
    carrier_name: str,
    tracking_number: str,
    tracking_url: str = "",
) -> str:
    """Build the localized customer note used by the legacy shipping endpoint."""
    text = CUSTOMER_NOTE_TEXT[customer_language(country)]
    carrier = html.escape(str(carrier_name or "Carrier"))
    tracking = html.escape(str(tracking_number or ""))
    if tracking_url:
        safe_url = html.escape(str(tracking_url), quote=True)
        tracking = f"<a href='{safe_url}'>{tracking}</a>"
    return text["shipped"].format(carrier=carrier, tracking=tracking)
