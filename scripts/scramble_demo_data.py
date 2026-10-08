"""
One-off anonymizer for internal-presentation copies of the client data files.

Deterministically replaces every company name, person/contact name, and email
address across the data files with a fake-but-plausible equivalent, so the same
real entity always maps to the same fake entity everywhere it appears (companies,
contacts, presets, logs, CSV, free text). Business-type words in company names
(e.g. "Consulting", "IT Solutions") are preserved; only the distinctive/brand part
is swapped. .env and templates.json (the firm's own identity, not client data)
are left untouched.

Run once from the repo root: python scripts/scramble_demo_data.py
"""
from __future__ import annotations

import csv
import hashlib
import json
import random
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

PUBLIC_DOMAINS = {
    "gmail.com", "yahoo.com", "yahoo.com.sg", "yahoo.co.in", "yahoo.com.hk",
    "yahoo.co.uk", "yahoo.com.ph", "yahoo.com.my", "yahoo.com.cn", "hotmail.com",
    "outlook.com", "live.com", "msn.com", "icloud.com", "me.com", "qq.com",
    "163.com", "126.com", "naver.com", "daum.net", "hanmail.net", "comcast.net",
    "btinternet.com", "aol.com", "protonmail.com", "gm.com", "mail.com",
    "sina.com", "foxmail.com", "rocketmail.com", "ymail.com",
}

LEGAL_SUFFIX_TOKENS = {
    "pte", "pte.", "ltd", "ltd.", "limited", "sdn", "sdn.", "bhd", "bhd.",
    "llp", "llc", "inc", "inc.", "co", "co.", "corp", "corp.", "corporation",
    "plc", "private", "pvt", "pvt.", "berhad",
}

DESCRIPTOR_KEYWORDS = {
    "consulting", "consultants", "consultancy", "solutions", "solution",
    "services", "service", "trading", "tradings", "logistics", "holdings",
    "holding", "group", "groups", "capital", "ventures", "venture",
    "investment", "investments", "investing", "management", "advisory",
    "advisors", "advisers", "labs", "lab", "laboratories", "laboratory",
    "technology", "technologies", "tech", "systems", "system",
    "manufacturing", "industries", "industry", "design", "designs",
    "studio", "studios", "media", "productions", "production", "marketing",
    "properties", "property", "realty", "engineering", "construction",
    "pharma", "pharmaceutical", "pharmaceuticals", "biotech",
    "biotechnology", "fashion", "apparel", "food", "foods", "beverage",
    "beverages", "shipping", "freight", "cargo", "retail", "digital",
    "creative", "creatives", "global", "worldwide", "network", "networks",
    "associates", "associate", "agency", "agencies", "enterprise",
    "enterprises", "international", "partners", "partnership",
    "partnerships", "kitchen", "restaurant", "restaurants", "cafe",
    "cafes", "hotel", "hotels", "tours", "tour", "travel", "travels",
    "operator", "operators", "tyre", "tyres", "specialty", "specialties",
    "products", "product", "supply", "supplies", "distribution",
    "distributors", "distributor", "chemical", "chemicals", "oil", "gas",
    "mining", "agriculture", "agro", "farm", "farms", "seafood", "bakery",
    "catering", "events", "event", "printing", "publishing",
    "entertainment", "games", "gaming", "toys", "jewellery", "jewelry",
    "watch", "watches", "cosmetics", "skincare", "beauty", "pharmacy",
    "research", "institute", "foundation", "club", "society", "store",
    "stores", "shop", "mart", "market", "markets", "supermarket",
    "boutique", "outlet", "gallery", "corporation", "equity", "fund",
    "funds", "asset", "assets", "wealth", "law", "legal", "accounting",
    "audit", "auditing", "tax", "finance", "financial", "insurance",
    "bank", "banking", "electronics", "electrical", "mechanical",
    "automotive", "auto", "motor", "motors", "warehouse", "interior",
    "architecture", "software", "hardware", "healthcare", "medical",
    "clinic", "clinics", "dental", "spa", "salon", "fitness", "sports",
    "education", "academy", "school", "training", "energy", "marine",
    "aviation", "security", "cleaning", "maintenance", "repair",
    "repairs", "rental", "rentals", "leasing", "florist", "florists",
    "flowers", "gifts", "gift", "packaging", "textile", "textiles",
    "garment", "garments", "furniture", "plumbing", "renovation",
    "contractors", "contractor", "communications", "telecom",
    "telecommunications", "resources", "resource", "machinery",
    "equipment", "tools", "tool", "chicken", "f&b",
}

BRAND_WORD_POOL = [
    "Bob", "Frank", "Alex", "Nova", "Orion", "Summit", "Vertex", "Crestline",
    "Meridian", "Apex", "Silverline", "Evergreen", "Northgate", "Lionsgate",
    "Starling", "Ironwood", "Clearpath", "Brightside", "Falcon", "Horizon",
    "Anchor", "Compass", "Pinecrest", "Ridgeline", "Westfield", "Eastgate",
    "Riverside", "Maple", "Cedar", "Granite", "Beacon", "Marcus", "Wei",
    "Hui", "Kiran", "Aisha", "Daniel", "Grace", "Leo", "Priya", "Hassan",
    "Mei", "Ravi", "Jasmine", "Victor", "Elena", "Samuel", "Nadia", "Felix",
    "Clara", "Omar", "Ling", "Arjun", "Nora", "Gavin", "Harper", "Dexter",
    "Mira", "Jonas", "Wendy", "Theo", "Sienna", "Bruno", "Keisha", "Milo",
    "Tessa",
]

DESCRIPTOR_WORD_POOL = [
    "Trading", "Consulting", "Holdings", "Enterprises", "Services",
    "Solutions", "Ventures", "Marketing", "Logistics", "Management",
]

FAKE_SUFFIXES = ["Pte Ltd", "Pte. Ltd.", "Sdn Bhd", "LLP"]

FIRST_NAME_POOL = [
    "James", "Mary", "Wei Ming", "Hui Ling", "Ahmad", "Siti", "Raj", "Priya",
    "Daniel", "Grace", "Kevin", "Michelle", "Hafiz", "Aisyah", "Kumar",
    "Anita", "Jun Jie", "Mei Hua", "David", "Laura", "Steven", "Rachel",
    "Suresh", "Lakshmi", "Farid", "Nurul", "Benjamin", "Charlotte",
    "Zhi Hao", "Xin Yi", "Arjun", "Meera", "Timothy", "Natalie", "Hakim",
    "Zainab", "Marcus", "Jasmine", "Vikram", "Divya", "Jonathan", "Sarah",
    "Kai Xin", "Li Wei", "Gavin", "Priyanka",
]

SURNAME_POOL = [
    "Tan", "Lim", "Lee", "Ng", "Wong", "Chen", "Goh", "Teo", "Ong", "Koh",
    "Chua", "Yeo", "Low", "Toh", "Chan", "Sim", "Loh", "Heng", "Seah",
    "Foo", "Abdullah", "Rahman", "Ibrahim", "Hassan", "Kumar", "Nair",
    "Pillai", "Singh", "Sharma", "Gupta", "Smith", "Johnson", "Williams",
    "Brown", "Taylor", "Wilson", "Anderson", "Thomas", "Martin",
]

GENERIC_LOCAL_PARTS = {
    "info", "admin", "administrator", "sales", "contact", "support",
    "enquiry", "enquiries", "hello", "accounts", "finance", "hr", "office",
    "mail", "general", "customerservice", "cs", "service", "services",
    "team", "marketing", "biz", "operations", "ops", "billing",
}


def rng_for(key: str) -> random.Random:
    h = hashlib.sha256(key.encode("utf-8")).hexdigest()
    return random.Random(int(h[:16], 16))


def match_case(original: str, generated: str) -> str:
    letters = [c for c in original if c.isalpha()]
    if letters and all(c.isupper() for c in letters):
        return generated.upper()
    if letters and all(c.islower() for c in letters):
        return generated.lower()
    return generated


def looks_like_company(name: str) -> bool:
    for tok in name.split():
        if tok.lower().strip(".,") in LEGAL_SUFFIX_TOKENS:
            return True
    return False


class Anonymizer:
    def __init__(self) -> None:
        self.company_cache: dict[str, str] = {}
        self.person_cache: dict[str, str] = {}
        self.domain_cache: dict[str, str] = {}
        self.email_cache: dict[str, str] = {}

        self.used_companies: set[str] = set()
        self.used_persons: set[str] = set()
        self.used_domains: set[str] = set()
        self.used_emails: set[str] = set()

        self.email_to_company: dict[str, str] = {}
        self.email_to_person: dict[str, str] = {}
        self.domain_to_company: dict[str, str] = {}

        self._test_counter = 0
        self._fallback_counter = 0

    # ---- association pass -------------------------------------------------

    def note_company(self, email: str | None, company: str | None) -> None:
        if not email or not company:
            return
        e = email.strip().strip("'\"").lower()
        c = company.strip()
        if e and c and e not in self.email_to_company:
            self.email_to_company[e] = c

    def note_person(self, email: str | None, name: str | None) -> None:
        if not email or not name:
            return
        e = email.strip().strip("'\"").lower()
        n = name.strip()
        if not e or not n or looks_like_company(n):
            return
        if n.replace(" ", "").replace(".", "").isdigit():
            return
        if e not in self.email_to_person:
            self.email_to_person[e] = n

    def finalize_domain_map(self) -> None:
        for e, c in self.email_to_company.items():
            if "@" not in e:
                continue
            domain = e.rsplit("@", 1)[1]
            if domain not in PUBLIC_DOMAINS and domain not in self.domain_to_company:
                self.domain_to_company[domain] = c

    # ---- ensure uniqueness --------------------------------------------------

    def _unique(self, candidate: str, used: set[str], is_email: bool = False) -> str:
        if candidate not in used:
            used.add(candidate)
            return candidate
        i = 2
        while True:
            if is_email:
                local, domain = candidate.split("@", 1)
                trial = f"{local}{i}@{domain}"
            else:
                trial = f"{candidate} {i}"
            if trial not in used:
                used.add(trial)
                return trial
            i += 1

    def _unique_domain(self, domain: str) -> str:
        if domain not in self.used_domains:
            self.used_domains.add(domain)
            return domain
        slug, _, tld = domain.rpartition(".")
        i = 2
        while True:
            trial = f"{slug}{i}.{tld}"
            if trial not in self.used_domains:
                self.used_domains.add(trial)
                return trial
            i += 1

    # ---- company name scrambling --------------------------------------------

    def scramble_company(self, original: str) -> str:
        key = original.strip().lower()
        if not key:
            return original
        if key in self.company_cache:
            return self.company_cache[key]

        tokens = original.strip().split()
        i = len(tokens)
        suffix_tokens: list[str] = []
        while i > 0 and tokens[i - 1].lower().strip(".,") in LEGAL_SUFFIX_TOKENS:
            suffix_tokens.insert(0, tokens[i - 1])
            i -= 1
        core_tokens = tokens[:i]

        j = len(core_tokens)
        desc_tokens: list[str] = []
        while j > 0 and core_tokens[j - 1].lower().strip(".,&()") in DESCRIPTOR_KEYWORDS:
            desc_tokens.insert(0, core_tokens[j - 1])
            j -= 1
        brand_tokens = core_tokens[:j]

        r = rng_for("COMPANY::" + key)

        if not desc_tokens:
            fake_brand = r.choice(BRAND_WORD_POOL)
            fake_desc = r.choice(DESCRIPTOR_WORD_POOL)
            fake_suffix = suffix_tokens if suffix_tokens else (
                [r.choice(FAKE_SUFFIXES)] if r.random() < 0.7 else []
            )
            result_tokens = [fake_brand, fake_desc] + fake_suffix
        else:
            fake_brand = r.choice(BRAND_WORD_POOL) if brand_tokens else r.choice(BRAND_WORD_POOL)
            result_tokens = [fake_brand] + desc_tokens + suffix_tokens

        result = " ".join(result_tokens)
        result = match_case(original, result)
        result = self._unique(result, self.used_companies)
        self.company_cache[key] = result
        return result

    # ---- person name scrambling ----------------------------------------------

    def scramble_person(self, original: str) -> str:
        key = original.strip().lower()
        if not key:
            return original
        if key in self.person_cache:
            return self.person_cache[key]

        if looks_like_company(original):
            result = self.scramble_company(original)
            self.person_cache[key] = result
            return result

        tokens = [t for t in original.split() if t]
        r = rng_for("PERSON::" + key)
        n = max(1, min(len(tokens), 4))
        fake_tokens = []
        for idx in range(n):
            fake_tokens.append(r.choice(FIRST_NAME_POOL) if idx == 0 else r.choice(SURNAME_POOL))
        result = " ".join(fake_tokens)
        result = match_case(original, result)
        result = self._unique(result, self.used_persons)
        self.person_cache[key] = result
        return result

    # ---- domain scrambling --------------------------------------------------

    def _slug(self, text: str) -> str:
        slug = re.sub(r"[^a-z0-9]", "", text.lower())
        return slug or "example"

    def fake_domain_for(self, domain: str) -> str:
        if domain in self.domain_cache:
            return self.domain_cache[domain]
        company = self.domain_to_company.get(domain)
        fake_company = self.scramble_company(company) if company else self.scramble_company(domain)
        tld = domain.split(".")[-1] if "." in domain else "com"
        if tld not in {"com", "net", "org", "sg", "my", "co", "hk", "io", "biz", "asia"}:
            tld = "com"
        result = f"{self._slug(fake_company)}.{tld}"
        result = self._unique_domain(result)
        self.domain_cache[domain] = result
        return result

    # ---- email scrambling ---------------------------------------------------

    def _local_from_name(self, fake_name: str) -> str:
        parts = [p.lower() for p in re.split(r"\s+", fake_name) if p.isalpha()]
        if len(parts) >= 2:
            return f"{parts[0]}.{parts[-1]}"
        if parts:
            return parts[0]
        return ""

    def _unique_test_email(self) -> str:
        while True:
            self._test_counter += 1
            candidate = f"testuser{self._test_counter:04d}@example-test.com"
            if candidate not in self.used_emails:
                self.used_emails.add(candidate)
                return candidate

    def scramble_email(self, original: str) -> str:
        key = original.strip().lower()
        if not key:
            return original
        if key in self.email_cache:
            return self.email_cache[key]

        if "@" not in key:
            result = self._unique_test_email()
            self.email_cache[key] = result
            return result

        local, domain = key.rsplit("@", 1)
        local_clean = local.strip("'\"")
        domain_clean = domain.strip("'\"")

        if domain_clean in PUBLIC_DOMAINS:
            person = self.email_to_person.get(key)
            company = self.email_to_company.get(key)
            if local_clean in GENERIC_LOCAL_PARTS and company:
                new_local = self._slug(self.scramble_company(company))[:16]
                new_domain = domain_clean
            elif person:
                new_local = self._local_from_name(self.scramble_person(person)) or local_clean
                new_domain = domain_clean
            elif company:
                new_local = self._slug(self.scramble_company(company))[:16]
                new_domain = domain_clean
            else:
                result = self._unique_test_email()
                self.email_cache[key] = result
                return result
        else:
            new_domain = self.fake_domain_for(domain_clean)
            if local_clean in GENERIC_LOCAL_PARTS:
                new_local = local_clean
            else:
                person = self.email_to_person.get(key)
                if person:
                    new_local = self._local_from_name(self.scramble_person(person))
                else:
                    new_local = ""
                if not new_local:
                    self._fallback_counter += 1
                    new_local = f"user{self._fallback_counter}"

        result = self._unique(f"{new_local}@{new_domain}", self.used_emails, is_email=True)
        self.email_cache[key] = result
        return result


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def dump_json(path: Path, data) -> None:
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def main() -> None:
    az = Anonymizer()

    companies = load_json(ROOT / "companies.json")
    person_names = load_json(ROOT / "person_names.json")
    profile_emails = load_json(ROOT / "profile_emails.json")
    left_out_companies = load_json(ROOT / "left_out_companies.json")
    subscriptions = load_json(ROOT / "subscriptions.json")
    presets_files = {
        name: load_json(ROOT / name)
        for name in [
            "presets.json",
            "presets.backup.json",
            "presets.before-left-out.json",
            "presets.before-sent-backfill.json",
            "presets.before-zoho.json",
        ]
    }
    profile_emails_files = {
        name: load_json(ROOT / name)
        for name in [
            "profile_emails.before-left-out.json",
            "profile_emails.before-sent-backfill.json",
        ]
    }
    zoho_contacts = load_json(ROOT / "zoho_contacts.json")

    change_log_lines = (ROOT / "change_log.jsonl").read_text(encoding="utf-8").splitlines()
    change_log = [json.loads(line) for line in change_log_lines if line.strip()]

    with open(ROOT / "drafts_log.csv", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        drafts_fieldnames = reader.fieldnames
        drafts_rows = list(reader)

    left_out_lines = (ROOT / "left_out.txt").read_text(encoding="utf-8").splitlines()
    import_conflicts_lines = (ROOT / "import_conflicts.txt").read_text(encoding="utf-8").splitlines()

    # ---- pass 1: build associations ----------------------------------------
    for item in companies:
        az.note_company(item.get("email"), item.get("company_name"))
    for email, name in person_names.items():
        az.note_person(email, name)
    for company, emails in profile_emails.items():
        for e in emails:
            az.note_company(e, company)
    for item in left_out_companies:
        for e in item.get("emails", []):
            az.note_company(e, item.get("company_name"))
    for items in presets_files.values():
        for item in items:
            az.note_company(item.get("email"), item.get("company_name"))
    for item in zoho_contacts:
        if item.get("company_name"):
            az.note_company(item.get("email"), item.get("company_name"))
        if item.get("name"):
            az.note_person(item.get("email"), item.get("name"))
    for entry in change_log:
        az.note_company(entry.get("email"), entry.get("company"))
    for row in drafts_rows:
        az.note_company(row.get("to"), row.get("client_name"))

    az.finalize_domain_map()

    # ---- pass 2: rewrite everything -----------------------------------------
    for item in companies:
        if item.get("company_name"):
            item["company_name"] = az.scramble_company(item["company_name"])
        if item.get("email"):
            item["email"] = az.scramble_email(item["email"])
    dump_json(ROOT / "companies.json", companies)

    new_person_names = {}
    for email, name in person_names.items():
        new_person_names[az.scramble_email(email)] = az.scramble_person(name)
    dump_json(ROOT / "person_names.json", new_person_names)

    new_profile_emails = {}
    for company, emails in profile_emails.items():
        new_company_key = az.scramble_company(company).lower()
        new_profile_emails[new_company_key] = [az.scramble_email(e) for e in emails]
    dump_json(ROOT / "profile_emails.json", new_profile_emails)

    for item in left_out_companies:
        if item.get("company_name"):
            item["company_name"] = az.scramble_company(item["company_name"])
        item["emails"] = [az.scramble_email(e) for e in item.get("emails", [])]
    dump_json(ROOT / "left_out_companies.json", left_out_companies)

    new_subscriptions = {}
    for key, rec in subscriptions.items():
        if rec.get("company_name"):
            rec["company_name"] = az.scramble_company(rec["company_name"])
        new_key = rec["company_name"].lower() if rec.get("company_name") else az.scramble_company(key).lower()
        new_subscriptions[new_key] = rec
    dump_json(ROOT / "subscriptions.json", new_subscriptions)

    for fname, items in presets_files.items():
        for item in items:
            if item.get("company_name"):
                item["company_name"] = az.scramble_company(item["company_name"])
            if item.get("email"):
                item["email"] = az.scramble_email(item["email"])
        dump_json(ROOT / fname, items)

    for fname, mapping in profile_emails_files.items():
        new_mapping = {}
        for company, emails in mapping.items():
            new_company_key = az.scramble_company(company).lower()
            new_mapping[new_company_key] = [az.scramble_email(e) for e in emails]
        dump_json(ROOT / fname, new_mapping)

    for item in zoho_contacts:
        if item.get("company_name"):
            item["company_name"] = az.scramble_company(item["company_name"])
        if item.get("name"):
            item["name"] = az.scramble_person(item["name"])
        if item.get("email"):
            item["email"] = az.scramble_email(item["email"])
    dump_json(ROOT / "zoho_contacts.json", zoho_contacts)

    new_change_log_lines = []
    for entry in change_log:
        if entry.get("company"):
            entry["company"] = az.scramble_company(entry["company"])
        if entry.get("email"):
            entry["email"] = az.scramble_email(entry["email"])
        new_change_log_lines.append(json.dumps(entry, ensure_ascii=False))
    (ROOT / "change_log.jsonl").write_text("\n".join(new_change_log_lines) + "\n", encoding="utf-8")

    for row in drafts_rows:
        old_client = row.get("client_name") or ""
        new_client = az.scramble_company(old_client) if old_client else old_client
        if row.get("subject") and old_client:
            row["subject"] = row["subject"].replace(old_client, new_client)
        row["client_name"] = new_client
        if row.get("to"):
            row["to"] = az.scramble_email(row["to"])
    with open(ROOT / "drafts_log.csv", "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=drafts_fieldnames)
        writer.writeheader()
        writer.writerows(drafts_rows)

    left_out_pattern = re.compile(r"^(left out \(\d+ emails?, rule not met\): )(.+)$")
    new_left_out_lines = []
    for line in left_out_lines:
        m = left_out_pattern.match(line)
        if m:
            new_left_out_lines.append(m.group(1) + az.scramble_company(m.group(2)))
        else:
            new_left_out_lines.append(line)
    (ROOT / "left_out.txt").write_text("\n".join(new_left_out_lines) + "\n", encoding="utf-8")

    email_pattern = re.compile(r"[A-Za-z0-9_.+\-']+@[A-Za-z0-9\-.]+\.[A-Za-z]{2,}")
    new_import_conflicts_lines = []
    for line in import_conflicts_lines:
        prefix_match = re.match(r"^([^/]+)( / .*)$", line)
        if prefix_match:
            label = prefix_match.group(1).strip()
            rest = prefix_match.group(2)
            if label.lower() in az.company_cache:
                line = az.company_cache[label.lower()] + rest
        line = email_pattern.sub(lambda m: az.scramble_email(m.group(0)), line)
        new_import_conflicts_lines.append(line)
    (ROOT / "import_conflicts.txt").write_text("\n".join(new_import_conflicts_lines) + "\n", encoding="utf-8")

    print("Companies scrambled:", len(az.company_cache))
    print("Persons scrambled:", len(az.person_cache))
    print("Emails scrambled:", len(az.email_cache))

    # copy scrambled seed files into the build_seed/ and dist/ seed mirrors
    seed_files = [
        "companies.json", "drafts_log.csv", "left_out_companies.json",
        "person_names.json", "presets.json", "profile_emails.json",
        "subscriptions.json",
    ]
    seed_dirs = [ROOT / "build_seed", ROOT / "dist" / "ClientEmailReminders" / "_internal" / "seed"]
    for d in seed_dirs:
        for fname in seed_files:
            src = ROOT / fname
            dst = d / fname
            if dst.exists():
                dst.write_bytes(src.read_bytes())


if __name__ == "__main__":
    main()
