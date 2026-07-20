"""
frontend/config.py
Platform configuration — colours, labels, file settings, download filenames.
Import this anywhere a platform constant is needed.
"""

PLATFORMS: dict = {
    "shopee": {
        "label":      "Shopee",
        "tab_icon":   "🛍️",
        "color":      "#1E5799",   # primary (header, button)
        "light":      "#E6F1FB",   # badge background
        "dark_text":  "#0C447C",   # badge text
        "file_types": ["xlsx"],
        "file_hint":  "Upload BOTH per month: Overview (Visitors) + Traffic (Orders). "
                      "CVR = Orders ÷ Visitors. Keep the month in the filename, "
                      "e.g. 'Shopee Traffic_January_2026.xlsx'.",
        "dl_name":    "Shopee_CVR_Traffic_Conversion.xlsx",
    },
    "tiktok": {
        "label":      "TikTok",
        "tab_icon":   "🎵",
        "color":      "#CC5500",
        "light":      "#FAE5D3",
        "dark_text":  "#8B3300",
        "file_types": ["xlsx"],
        "file_hint":  "TikTok Shop Analytics Key Metrics exports (.xlsx)",
        "dl_name":    "TikTok_CVR_Traffic_Conversion.xlsx",
    },
    "lazada": {
        "label":      "Lazada",
        "tab_icon":   "🛒",
        "color":      "#0F6E3A",
        "light":      "#D6EDDF",
        "dark_text":  "#0A4D28",
        "file_types": ["xls", "xlsx"],
        "file_hint":  "Lazada Business Advisor Dashboard exports (.xls or .xlsx)",
        "dl_name":    "Lazada_CVR_Traffic_Conversion.xlsx",
    },
}
