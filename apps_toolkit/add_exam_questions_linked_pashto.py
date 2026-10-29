import os

PASHTO_DIR = r"H:\My Drive\PMS KPK 2022 notes\4. Pashto"

updates = {
    "Q1_and_Q2_Pashto_Compulsory_Translations_and_Proverbs_2025_Master_Notes.md": """### 🎯 EXAM QUESTIONS LINKED (د تېرو کلونو پوښتنې په پوره متن سره)

> • **2025 Q.1 (الف)**: "A Clean and Green Pakistan is not merely an aesthetic ambition-it is the lifeline of our nation's future. The land that once echoed with the fragrance of orange blossoms and jasmine now struggles under the weight of urban smog and deforestation. As the adage goes, “We do not inherit the earth from our ancestors; we borrow it from our children.” If we continue to treat our soil as a dumping ground and our rivers as sewers, we will sow the wind and reap the whirlwind. The road to a verdant Pakistan lies in grassroots action-community gardens in barren plots, rooftop plantations in crowded cities, and eco-friendly alternatives replacing harmful plastics. When citizens become custodians of their environment, cleanliness transforms from a campaign into a culture, and Pakistan can blossom into the emerald heart of South Asia once more." (14 Marks)
>
> • **2025 Q.1 (ب)**: "Translate any THREE of the following proverbs into authentic Pashto: 1. When the head is alive, caps are many. 2. By the time truth is known, lies will have destroyed the village. 3. If you have money, there is no need to make request. 4. When you reach old age you become pious. 5. As the donkey, so the load. 6. Some say, what shall I do with it, others say where shall I hide it." (6 Marks)
>
> • **2025 Q.2 (الف)**: "خداے پاک چې په یو لوشي قوم داسې لازم الر او بیاتي په کښې د عبادت فخر ورکړه، کرک او هم دستې د ګلونو ورمۍ، د خلوص خوند، ښۀ اخلاق، صبر او قربانۍ ثمر په کښې شته کول او بیاتي خپل لوړے حکمت سره د بښنې او درګزر په یو انساني صورت کښې تماره کړه او د مور په شکل کښې ئې دا دنیا مخ ته راواستوله، هر ساربان وائي؛ مور آسمان نه راښکته شوے دعا ده، د انسان ژوند د تولولو لوئے، غور، ګران، په زړۀ پورې او ښکلې لفظ ،مور، دے." (20 Marks)
>
> • **2022 Q.2**: "دغه لاندینی عبارت د انګرېزي نه پښتو کښې ترجمه کړئ او په متلونو (Proverbs) کښې درې د انګرېزي نه پښتو ته واړوئ۔" (20 Marks)

---
""",

    "Q3_Abdul_Hamid_Baba_Poetic_Art_and_Imagery_2025_Master_Notes.md": """### 🎯 EXAM QUESTIONS LINKED (د تېرو کلونو پوښتنې په پوره متن سره)

> • **2025 Q.3**: "د عبدالحميد بابا شاعرۍ د تخییل، صنائع، استعارو او تشبیهاتو په لحاظ ځانګړې معیار لري، د دې نکتو په رڼا کښې د هغۀ په شاعرۍ بحث وکړئ۔" (20 Marks)
>
> • **2022 Q.3**: "د عبدالحميد بابا شاعرۍ د تخییل، صنائع، استعارو او تشبیهاتو په لحاظ ځانګړی معیار لري، د دې نکتو په رڼا کې د هغۀ په شاعرۍ بحث وکړئ۔" (20 Marks)
>
> • **2022 Q.1 (MCQ #4)**: "پښتو کښې موشګاف شاعر چا ته وئیلے شی؟ (حمید مومند)" (1 Mark)

---
""",

    "Q4_Pashto_Folk_Literature_Folklore_Genres_2025_Master_Notes.md": """### 🎯 EXAM QUESTIONS LINKED (د تېرو کلونو پوښتنې په پوره متن سره)

> • **2025 Q.4**: "د پښتو اولسي ادب چي زمونږ ژبي او ژوند امین او ضامن دے په قیصو، متلونو، ټپو، چاربیتو او بدلو باندې اړخ لري، په دې اصنافو باندې تفصیلي رڼا واچوئ۔" (20 Marks)
>
> • **2022 Q.4**: "د پښتو اولسي ادب چې زموږ ژبې او ژوند امین او ضامن دی په قیصو، متلونو، ټپو، چاربیتو او بدلو باندې اړخ لري، په دې اصنافو باندې تفصیلي رڼا واچوئ۔" (20 Marks)
>
> • **2022 Q.1 (MCQ #11)**: "فولکلور څه ته وائي؟ (اولسي ادب)" (1 Mark)

---
""",

    "Q5_Evolution_of_Pashto_Drama_Radio_Stage_TV_2025_Master_Notes.md": """### 🎯 EXAM QUESTIONS LINKED (د تېرو کلونو پوښتنې په پوره متن سره)

> • **2025 Q.5**: "د پښتو ډرامې په ابتدا او ارتقاء خپلو خیالاتو څرګندونه کړئ۔" (20 Marks)
>
> • **2022 Q.5**: "د پښتو ډرامې په ابتدا او ارتقاء خپلو خیالاتو څرګندونه کړئ۔" (20 Marks)
>
> • **2022 Q.1 (MCQ #14)**: "دری یتیمان ډرامه چا لیکلی دہ؟ (عبدالاکبر خان اکبر)" (1 Mark)
>
> • **2022 Q.1 (MCQ #19)**: "پښتو ادب کښې کردارنګاري د چا د پی ایچ ډی مقاله دہ؟ (ډاکټر محمد اعظم اعظم)" (1 Mark)

---
""",

    "Q6_Rahman_Baba_Poetic_Art_Aesthetics_and_Sufism_2025_Master_Notes.md": """### 🎯 EXAM QUESTIONS LINKED (د تېرو کلونو پوښتنې په پوره متن سره)

> • **2025 Q.6**: "د رحمان بابا د کلام په فني خصوصیاتو او شعري ښکلا باندې بحث وکړئ او په مثالونو کښې ئې شعرونه خویش کړئ۔" (20 Marks)
>
> • **2022 Q.6**: "د رحمان بابا د کلام په فني خصوصیاتو او شعري ښکلا باندې بحث وکړئ او په مثالونو کې یې شعرونه خویش کړئ۔" (20 Marks)
>
> • **2022 Q.1 (MCQ #3)**: "د رحمان بابا د پلار نوم څه ؤ؟ (عبدالستار)" (1 Mark)

---
""",

    "Q7_Modern_Pashto_Poets_Sayel_Salim_Raz_Azam_2025_Master_Notes.md": """### 🎯 EXAM QUESTIONS LINKED (د تېرو کلونو پوښتنې په پوره متن سره)

> • **2025 Q.7**: "په لاندې ورکړې شوې جديدو شاعرانو کښې د یو شاعر په ژوند او فن باندې خپله مطالعه تجزيه او ليکئ: ۱. سلیم راز، ۲. داکتر محمد اعظم اعظم، ۳. رحمت شاه سائل۔" (20 Marks)
>
> • **2022 Q.7**: "په لاندې ورکړې شویو جدیدو شاعرانو کې د یو شاعر په ژوند او فن باندې خپله مطالعه تجزیه او ولیکئ: ۱. سلیم راز، ۲. ډاکټر محمد اعظم اعظم، ۳. رحمت شاه سائل۔" (20 Marks)

---
"""
}

def apply_updates():
    for filename, block in updates.items():
        file_path = os.path.join(PASHTO_DIR, filename)
        if not os.path.exists(file_path):
            continue
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
        
        target = "## 🎯 د ازموینې کلیدي نکتې (Executive Exam Takeaways)"
        if "### 🎯 EXAM QUESTIONS LINKED" not in content and target in content:
            new_content = content.replace(target, block + "\n" + target)
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(new_content)
            print(f"Updated: {filename}")
        else:
            print(f"Skipped or already present: {filename}")

if __name__ == "__main__":
    apply_updates()
