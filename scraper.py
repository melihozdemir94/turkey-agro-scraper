import os
import requests
import pandas as pd

# Türkiye'nin 7 Coğrafi Bölgesi ve İlleri
REGIONS = {
    "Marmara": ["İstanbul", "Bursa", "Balıkesir", "Kocaeli", "Tekirdağ", "Sakarya", "Çanakkale", "Edirne", "Kırklareli", "Yalova", "Bilecik"],
    "Ic_Anadolu": ["Ankara", "Konya", "Kayseri", "Eskişehir", "Sivas", "Kırıkkale", "Aksaray", "Karaman", "Kırşehir", "Niğde", "Nevşehir", "Yozgat", "Çankırı"],
    "Ege": ["İzmir", "Manisa", "Aydın", "Denizli", "Muğla", "Afyonkarahisar", "Kütahya", "Uşak"],
    "Akdeniz": ["Antalya", "Adana", "Mersin", "Hatay", "Kahramanmaraş", "Isparta", "Burdur", "Osmaniye"],
    "Karadeniz": ["Samsun", "Trabzon", "Ordu", "Tokat", "Zonguldak", "Çorum", "Kastamonu", "Rize", "Artvin", "Sinop", "Bartın", "Amasya", "Giresun", "Gümüşhane", "Karabük", "Düzce"],
    "Dogu_Anadolu": ["Erzurum", "Van", "Malatya", "Elazığ", "Erzincan", "Ağrı", "Kars", "Bitlis", "Muş", "Iğdır", "Bingöl", "Hakkari", "Tunceli", "Ardahan"],
    "Guneydogu_Anadolu": ["Gaziantep", "Diyarbakır", "Şanlıurfa", "Mardin", "Adıyaman", "Batman", "Siirt", "Kilis", "Şırnak"]
}

KEYWORDS = ["tarım", "gübre", "damla sulama", "ziraat"]

def fetch_osm_data(city, keyword):
    url = "https://overpass-api.de/api/interpreter"
    query = f"""
    [out:json][timeout:25];
    area[name="{city}"]->.searchArea;
    (
      node["shop"~"agrarian|farm|hardware"](area.searchArea);
      node["name"~"{keyword}", i](area.searchArea);
      way["name"~"{keyword}", i](area.searchArea);
    );
    out body;
    >;
    out skel qt;
    """
    try:
        response = requests.post(url, data={'data': query}, timeout=30)
        if response.status_code == 200:
            data = response.json()
            elements = data.get("elements", [])
            results = []
            for el in elements:
                tags = el.get("tags", {})
                if "name" in tags:
                    results.append({
                        "Firma İsmi": tags.get("name"),
                        "İlçe/Konum": tags.get("addr:suburb", tags.get("addr:district", "Merkez")),
                        "Açık Adres": tags.get("addr:street", "-"),
                        "Telefon": tags.get("phone", tags.get("contact:phone", "-")),
                        "Web Sitesi": tags.get("website", tags.get("contact:website", "-")),
                        "E-posta": tags.get("email", "-")
                    })
            return results
    except Exception as e:
        print(f"Hata ({city} - {keyword}): {e}")
    return []

def main():
    for region_name, cities in REGIONS.items():
        output_file = f"{region_name}_Tarim_Bayileri.xlsx"
        print(f"[{region_name}] Bölgesi taranıyor...")
        
        all_city_dfs = []
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for city in cities:
                city_data = []
                for kw in KEYWORDS:
                    res = fetch_osm_data(city, kw)
                    city_data.extend(res)
                
                if city_data:
                    df = pd.DataFrame(city_data).drop_duplicates(subset=["Firma İsmi"])
                else:
                    df = pd.DataFrame(columns=["Firma İsmi", "İlçe/Konum", "Açık Adres", "Telefon", "Web Sitesi", "E-posta"])
                
                sheet_name = city[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
        print(f"[{region_name}] Tamamlandı -> {output_file} kaydedildi.\n")

if __name__ == "__main__":
    main()
