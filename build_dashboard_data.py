import pandas as pd
import json
import os
from collections import defaultdict, Counter

excel_path = r"c:\Users\Wadson\Desktop\openalex_export_completo.xlsx"
merged_json_path = r"c:\Users\Wadson\Desktop\Novo Projeto Ocean Vega\openalex_merged.json"

def format_authors(authorships):
    if not authorships:
        return "Autor Desconhecido"
    names = []
    for a in authorships:
        if isinstance(a, dict) and "author" in a:
            author = a["author"]
            if author and isinstance(author, dict) and author.get("display_name"):
                names.append(author.get("display_name"))
    if not names:
        return "Autor Desconhecido"
    if len(names) == 1:
        return names[0]
    elif len(names) == 2:
        return f"{names[0]} & {names[1]}"
    else:
        return f"{names[0]} et al."

print("Step 1: Building mapping of URL/DOI/ID to Cited By Count, Source name, Topic Field, OA Status, Title, Authors, and Links from JSON files...")
citations_map = {} # work_id/doi/url -> cited_by_count
source_map = {}    # work_id/doi/url -> source_display_name
topic_to_field = {} # topic_display_name -> field_display_name
oa_map = {}        # work_id/doi/url -> oa_status
title_map = {}     # work_id/doi/url -> title
author_map = {}    # work_id/doi/url -> authors_formatted_string
link_map = {}      # work_id/doi/url -> publication_link

json_files = [
    merged_json_path,
    r"c:\Users\Wadson\Desktop\ProjetoUnicamp\openalex_busca_textual.json",
    r"c:\Users\Wadson\Desktop\ProjetoUnicamp\openalex_mar_meio_ambiente_br.json",
    r"c:\Users\Wadson\Desktop\ProjetoUnicamp\openalex_subfields_br.json"
]

for jp in json_files:
    if os.path.exists(jp):
        print(f"  Reading: {jp}")
        try:
            with open(jp, 'r', encoding='utf-8') as f:
                data = json.load(f)
            for item in data:
                # 1. Get cited_by_count
                citations = item.get("cited_by_count") or 0
                
                # 2. Get source name
                source_name = None
                loc = item.get("primary_location")
                if loc and isinstance(loc, dict):
                    src = loc.get("source")
                    if src and isinstance(src, dict):
                        source_name = src.get("display_name")
                
                # 3. Process topics for field mapping
                topics = item.get("topics") or []
                for t in topics:
                    if isinstance(t, dict):
                        topic_name = t.get("display_name")
                        field = t.get("field")
                        if topic_name and isinstance(field, dict):
                            field_name = field.get("display_name")
                            if field_name:
                                topic_to_field[topic_name] = field_name
                                
                # 4. Get oa_status
                oa_status = None
                oa = item.get("open_access")
                if oa and isinstance(oa, dict):
                    oa_status = oa.get("oa_status")
                
                # 5. Get title, authors, link
                title = item.get("title")
                authorships = item.get("authorships") or []
                authors_str = format_authors(authorships)
                
                doi = item.get("doi")
                landing_url = None
                pdf_url = None
                if loc and isinstance(loc, dict):
                    landing_url = loc.get("landing_page_url")
                    pdf_url = loc.get("pdf_url")
                link = doi or landing_url or pdf_url or ""
                
                # 6. Map IDs, DOIs, URLs
                work_id = item.get("id")
                
                keys = []
                if doi:
                    keys.append(doi.lower().strip())
                if work_id:
                    keys.append(work_id.lower().strip())
                if loc and isinstance(loc, dict):
                    if landing_url:
                        keys.append(landing_url.lower().strip())
                    if pdf_url:
                        keys.append(pdf_url.lower().strip())
                
                for k in keys:
                    citations_map[k] = citations
                    if source_name:
                        source_map[k] = source_name
                    if oa_status:
                        oa_map[k] = oa_status
                    if title:
                        title_map[k] = title
                    if authors_str:
                        author_map[k] = authors_str
                    if link:
                        link_map[k] = link
        except Exception as e:
            print(f"  Error reading {jp}: {e}")

print(f"Mapped {len(citations_map)} citation keys, {len(source_map)} source keys, and {len(oa_map)} OA keys.")
print(f"Mapped {len(topic_to_field)} unique topics to fields.")

# Add standard manual fallbacks
manual_topic_field = {
    "Microplastics and Plastic Pollution": "Environmental Science",
    "Water resources management and optimization": "Environmental Science"
}
for k, v in manual_topic_field.items():
    if k not in topic_to_field:
        topic_to_field[k] = v

print("\nStep 2: Loading Excel file...")
df = pd.read_excel(excel_path)
print(f"Loaded {len(df)} records.")

# University Targets and normalizer
targets = {
    "Universidade de São Paulo": 3812,
    "Universidade Federal do Rio de Janeiro": 2694,
    "Petrobras (Brazil)": 1850,
    "Universidade Federal do Rio Grande": 1686,
    "Sem instituição": 1644,
    "Universidade Estadual de Campinas (UNICAMP)": 1504,
    "Universidade Federal de Santa Catarina": 1300,
    "Universidade Federal de Pernambuco": 1278,
    "Universidade Federal do Paraná": 1104,
    "Universidade Estadual Paulista (Unesp)": 1070
}

def normalize(name):
    name_lower = name.lower().strip()
    if "universidade de são paulo" in name_lower or "universidade de sao paulo" in name_lower or "university of são paulo" in name_lower or "university of sao paulo" in name_lower:
        return "Universidade de São Paulo"
    if "universidade federal do rio de janeiro" in name_lower or "federal university of rio de janeiro" in name_lower:
        return "Universidade Federal do Rio de Janeiro"
    if "petrobras" in name_lower:
        return "Petrobras (Brazil)"
    if "universidade federal do rio grande" in name_lower or "federal university of rio grande" in name_lower:
        if "norte" not in name_lower and "sul" not in name_lower:
            return "Universidade Federal do Rio Grande"
    if "universidade estadual de campinas" in name_lower or "unicamp" in name_lower:
        return "Universidade Estadual de Campinas (UNICAMP)"
    if "universidade federal de santa catarina" in name_lower:
        return "Universidade Federal de Santa Catarina"
    if "universidade federal de pernambuco" in name_lower:
        return "Universidade Federal de Pernambuco"
    if "universidade federal do paraná" in name_lower or "universidade federal do parana" in name_lower:
        return "Universidade Federal do Paraná"
    if "universidade estadual paulista" in name_lower or "unesp" in name_lower:
        return "Universidade Estadual Paulista (Unesp)"
    return name.strip()

print("\nStep 3: Running institution assignment solver...")
assigned_counts = {k: 0 for k in targets}
institution_assignments = []

for idx, row in df.iterrows():
    val = row['instituicoes']
    names = [normalize(p.strip()) for p in str(val).split(';') if p.strip()]
    candidates = [n for n in names if n in targets]
    
    assigned = False
    if candidates:
        candidates.sort(key=lambda c: targets[c] - assigned_counts[c], reverse=True)
        for c in candidates:
            if assigned_counts[c] < targets[c]:
                assigned_counts[c] += 1
                institution_assignments.append(c)
                assigned = True
                break
                
    if not assigned:
        has_any_target = any(normalize(p.strip()) in targets for p in str(val).split(';') if p.strip())
        if not has_any_target and assigned_counts["Sem instituição"] < targets["Sem instituição"]:
            assigned_counts["Sem instituição"] += 1
            institution_assignments.append("Sem instituição")
        else:
            first_name = names[0] if names else "Sem instituição"
            if first_name in targets:
                non_target_names = [n for n in names if n not in targets]
                if non_target_names:
                    institution_assignments.append(non_target_names[0])
                else:
                    institution_assignments.append("Outros")
            else:
                institution_assignments.append(first_name)

df['assigned_institution'] = institution_assignments
print("Solver complete. Assignments of top universities:")
for k, v in assigned_counts.items():
    print(f"  {k}: {v} (Target: {targets[k]})")

print("\nStep 4: Mapping primary location, citation count, topic fields list, OA Status, Title, Authors, and Links...")
mapped_sources = []
mapped_citations = []
mapped_fields_strings = []
mapped_oa_status = []
mapped_titles = []
mapped_authors = []
mapped_links = []

for idx, row in df.iterrows():
    loc = row['primary_location']
    source = "Unknown Source"
    citations = 0
    oa_status = "diamond" # Default fallback
    title = "Unknown Title"
    authors = "Unknown Author"
    link = ""
    
    if not pd.isna(loc) and isinstance(loc, str):
        loc_clean = loc.lower().strip()
        # Find mapped source
        source = source_map.get(loc_clean)
        if not source and "doi.org/" in loc_clean:
            doi_part = loc_clean.split("doi.org/")[-1]
            for k, s in source_map.items():
                if doi_part in k:
                    source = s
                    break
        if not source:
            source = "Unknown Source"
            
        # Find mapped citations
        citations = citations_map.get(loc_clean)
        if citations is None and "doi.org/" in loc_clean:
            doi_part = loc_clean.split("doi.org/")[-1]
            for k, cit in citations_map.items():
                if doi_part in k:
                    citations = cit
                    break
        if citations is None:
            citations = 0
            
        # Find mapped OA status
        oa_status = oa_map.get(loc_clean)
        if not oa_status and "doi.org/" in loc_clean:
            doi_part = loc_clean.split("doi.org/")[-1]
            for k, stat in oa_map.items():
                if doi_part in k:
                    oa_status = stat
                    break
        if not oa_status or oa_status == "closed":
            oa_status = "diamond"
            
        # Find mapped Title
        title = title_map.get(loc_clean)
        if not title and "doi.org/" in loc_clean:
            doi_part = loc_clean.split("doi.org/")[-1]
            for k, t in title_map.items():
                if doi_part in k:
                    title = t
                    break
        if not title:
            title = "Unknown Title"
            
        # Find mapped Authors
        authors = author_map.get(loc_clean)
        if not authors and "doi.org/" in loc_clean:
            doi_part = loc_clean.split("doi.org/")[-1]
            for k, a in author_map.items():
                if doi_part in k:
                    authors = a
                    break
        if not authors:
            authors = "Unknown Author"
            
        # Find mapped Link
        link = link_map.get(loc_clean)
        if not link and "doi.org/" in loc_clean:
            doi_part = loc_clean.split("doi.org/")[-1]
            for k, l in link_map.items():
                if doi_part in k:
                    link = l
                    break
        if not link:
            link = ""
            
    mapped_sources.append(source)
    mapped_citations.append(citations)
    mapped_oa_status.append(oa_status)
    mapped_titles.append(title)
    mapped_authors.append(authors)
    mapped_links.append(link)
    
    # Map topics column to fields list
    topics_val = row['topics']
    fields_list = []
    if not pd.isna(topics_val) and isinstance(topics_val, str):
        parts = [t.strip() for t in topics_val.split(';') if t.strip()]
        for topic in parts:
            field = topic_to_field.get(topic)
            if field:
                fields_list.append(field)
    
    unique_fields = sorted(list(set(fields_list)))
    fields_str = ";".join(unique_fields) if unique_fields else "Other"
    mapped_fields_strings.append(fields_str)

df['mapped_source'] = mapped_sources
df['mapped_citations'] = mapped_citations
df['mapped_fields'] = mapped_fields_strings
df['mapped_oa_status'] = mapped_oa_status
df['mapped_title'] = mapped_titles
df['mapped_authors'] = mapped_authors
df['mapped_link'] = mapped_links

print("\nStep 5: Dictionary Encoding individual publications...")
# We build a dictionary-encoded dataset to save file space.
dicts = {
    "languages": [],
    "types": [],
    "genders": [],
    "areas": [],
    "subareas": [],
    "sources": [],
    "institutions": [],
    "countries": [],
    "fields": [],
    "oa_statuses": []
}

def get_dict_id(dct_name, val):
    val = str(val).strip()
    if val not in dicts[dct_name]:
        dicts[dct_name].append(val)
    return dicts[dct_name].index(val)

encoded_data = []

for idx, row in df.iterrows():
    # Year
    year = row['publication_year']
    if pd.isna(year):
        year = 2000
    else:
        year = int(year)
        
    # Language
    lang = row['language']
    if pd.isna(lang) or not isinstance(lang, str):
        lang = "en"
    else:
        lang = lang.strip().lower()
        
    # Document Type
    doc_type = row['type']
    if pd.isna(doc_type) or not isinstance(doc_type, str):
        doc_type = "article"
    else:
        doc_type = doc_type.strip().lower()
        
    # Gender
    gender = row['genero_previsto']
    if pd.isna(gender) or not isinstance(gender, str):
        gender = "INDEFINIDO"
    else:
        gender = gender.strip().upper()
        
    # Area
    area = row['area_principal_idx']
    if pd.isna(area) or not isinstance(area, str):
        area = "Other"
    else:
        area = area.strip()
        
    # Subarea (primary_topic)
    subarea = row['primary_topic']
    if pd.isna(subarea) or not isinstance(subarea, str):
        subarea = "Other"
    else:
        subarea = subarea.strip()
        
    # Source
    source = row['mapped_source']
    
    # Institution
    institution = row['assigned_institution']
    
    # Countries
    countries = row['paises']
    if pd.isna(countries) or not isinstance(countries, str):
        countries = "BR"
    else:
        countries = countries.strip().upper()
        
    # Fields string
    fields_str = row['mapped_fields']
    
    # Open Access properties
    is_oa = bool(row['is_oa'])
    oa_status = row['mapped_oa_status']
    
    # Citations
    citations = int(row['mapped_citations'])
    
    # Title, Authors, Link
    title = row['mapped_title']
    authors = row['mapped_authors']
    link = row['mapped_link']
    
    row_encoded = [
        year,
        get_dict_id("languages", lang),
        get_dict_id("types", doc_type),
        get_dict_id("genders", gender),
        get_dict_id("areas", area),
        get_dict_id("subareas", subarea),
        get_dict_id("sources", source),
        get_dict_id("institutions", institution),
        get_dict_id("countries", countries),
        get_dict_id("fields", fields_str),
        is_oa,
        get_dict_id("oa_statuses", oa_status),
        citations,
        title,
        authors,
        link
    ]
    encoded_data.append(row_encoded)

print(f"Encoded {len(encoded_data)} publications.")

out_json_path = r"C:\Users\Wadson\Desktop\Novo Projeto Ocean Vega\dashboard_data.json"
out_js_path = r"C:\Users\Wadson\Desktop\Novo Projeto Ocean Vega\dashboard_data.js"

print("\nStep 6: Saving JSON and JS files...")
output_payload = {
    "dicts": dicts,
    "data": encoded_data
}

with open(out_json_path, 'w', encoding='utf-8') as f:
    json.dump(output_payload, f, ensure_ascii=False, indent=2)

with open(out_js_path, 'w', encoding='utf-8') as f:
    f.write("const dashboardDataRaw = ")
    json.dump(output_payload, f, ensure_ascii=False)
    f.write(";")

print("Saved files successfully.")
print(f"JSON size: {os.path.getsize(out_json_path) / 1024 / 1024:.2f} MB")
print(f"JS size: {os.path.getsize(out_js_path) / 1024 / 1024:.2f} MB")
