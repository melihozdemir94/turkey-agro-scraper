import os
import requests
import pandas as pd
import time

# Türkiye'nin 7 Coğrafi Bölgesi ve Anahtar İlleri (Hızlı tarama için optimize edildi)
REGIONS = {
    "Marmara": ["İstanbul", "Bursa", "Kocaeli", "Sakarya", "Balıkesir"],
    "Ic_Anadolu": ["Ankara", "Konya", "Kayseri", "Eskişehir", "Sivas"],
    "Ege": ["İzmir", "Manisa", "Aydın", "Denizli", "Muğla"],
    "Akdeniz": ["Antalya", "Adana", "Mersin", "Hatay", "Kahramanmaraş"],
    "Karadeniz": ["Samsun", "Trabzon", "Ordu", "Tokat", "Zonguldak"],
    "Dogu_Anadolu": ["Erzurum", "Van", "Malatya", "Elazığ"],
    "Guneydogu_Anadolu": ["Gaziantep", "Diyarbakır", "Şanlıurfa", "Mardin"]
}

KEYWORDS = ["tarım", "gübre", "damla sulama", "ziraat"]

def fetch_osm_data(city, keyword):
    url = "https://overpass-api.de/api/interpreter"
    query = f"""
    [out:json][timeout:10];
    area[name="{city}"]->.searchArea;
    (
      node["shop"~"agrarian|farm"](area.searchArea);
      node["name"~"{keyword}", i](area.searchArea);
    );
    out body;
    """
    try:
        response = requests.post(url, data={'data': query}, timeout=15)
        if response.status_code == 200:
            data = response.json()
            elements = data.get("elements", [])
            results = []
            for el in elements:
                tags = el.get("tags", {})
                if "name" in tags:
                    results.append({
                        "Firma İsmi": tags.get("name"),
                        "İlçe/Konum": tags.get("addr:suburb", "Merkez"),
                        "Açık Adres": tags.get("addr:street", "-"),
                        "Telefon": tags.get("phone", "-"),
                        "Web Sitesi": tags.get("website", "-"),
                        "E-posta": tags.get("email", "-")
                    })
            return results
    except Exception as e:
        print(f"Atlandı ({city}): {e}")
    return []

def main():
    for region_name, cities in REGIONS.items():
        output_file = f"{region_name}_Tarim_Bayileri.xlsx"
        print(f"[{region_name}] taranıyor...")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for city in cities:
                city_data = []
                for kw in KEYWORDS:
                    res = fetch_osm_data(city, kw)
                    city_data.extend(res)
                    time.sleep(0.5) # Sunucuyu yormamak ve kilitlenmemek için kısa bekleme
                
                if city_data:
                    df = pd.DataFrame(city_data).drop_duplicates(subset=["Firma İsmi"])
                else:
                    df = pd.DataFrame(columns=["Firma İsmi", "İlçe/Konum", "Açık Adres", "Telefon", "Web Sitesi", "E-posta"])
                
                sheet_name = city[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
        print(f"[{region_name}] Tamamlandı -> {output_file}\n")

if __name__ == "__main__":
    main()
