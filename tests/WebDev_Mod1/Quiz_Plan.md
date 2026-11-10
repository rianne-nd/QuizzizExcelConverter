# Web Dev — Module 1 Quiz Plan

**Module:** Introduction to Mobile Application Development Fundamentals using Android
**Source:** 58 slides (Pages 1–58)
**Purpose of this document:** confirm that every part of the module will be covered before
any questions are written.

---

## 0. Final decisions (approved) — what was actually built

| Decision | Choice |
|---|---|
| Question count | **70** (exhaustive), not the 55 originally proposed |
| Android 13 | **Left out**, to match the module |
| Split | Written as one quiz first, **then split into 7 files of 10** |
| Downloads table | **Included** — as a Checkbox asking which two apps tied at 49M |
| Images | **None added.** The source has no image URLs, and inventing them was ruled out, so no `Image:` lines were used |

Delivered:

* `WebDev_Mod1Raw.txt` — all 70 questions
* `WebDev_Mod1_pt1.txt` … `WebDev_Mod1_pt7.txt` — 10 questions each
* `_QuizizzOutput\Quizizz_WebDev_Mod1.xlsx` — the full 70-question workbook
* `_QuizizzOutput\Quizizz_WebDev_Mod1_pt1.xlsx` … `_pt7.xlsx` — the seven 10-question workbooks

Final section counts: A 2 · B 9 · C 7 · D 6 · E 15 · F 28 · G 3 = **70**
Question types: 61 Multiple Choice, 9 Checkbox. All verified with 0 invalid answer keys.

---

## 1. Proposed quiz shape (original proposal, for reference)

| | |
|---|---|
| **Total questions** | 55 |
| **Question types** | Multiple Choice (one correct) and Checkbox (two or more correct) |
| **Options per question** | 4 for Multiple Choice, 4–5 for Checkbox |
| **Time per question** | 60 seconds |
| **Difficulty** | Mostly recall of what the module states, plus a few "which version introduced X" discriminations |
| **Output** | `WebDev_Mod1Raw.txt` → `Quizizz_WebDev_Mod1.xlsx`, ready for Wayground |

Checkbox questions are used where the module lists several things together (e.g. the seven
"Why Android?" reasons), because those facts are naturally "select all that apply".

---

## 2. Coverage map

Every content slide is assigned to a section below. The "Questions" column is what I intend to
write, so you can see nothing is over- or under-weighted.

### Section A — What is Mobile Application Development? (Pages 1–3) — **2 questions**

| Fact to test | Source |
|---|---|
| Definition: the act/process by which a mobile app is developed for mobile devices | p3 |
| Which devices count as "mobile" in the definition (PDAs, enterprise digital assistants) | p3 |

### Section B — What Android is (Pages 4–7) — **6 questions**

| Fact to test | Source |
|---|---|
| Android is an **open-source**, **Linux-based** OS for smartphones and tablets | p4 |
| Android is the world's best-selling smartphone platform | p5 |
| Apps are primarily written in **Java**, and mostly in **Kotlin** | p5 |
| The Android Market is Google's online app store (today the Google Play Store) | p5 |
| Of 3.553M Play Store apps, **3.442M (96.9%) are free** and ~110k (3.1%) are paid | p6 |
| Most-downloaded app of Nov 2023 by combined total (**TikTok, 49M**) | p6 |

### Section C — Why Android? (Pages 8–9) — **5 questions**

The module lists exactly seven reasons:

1. Open Source
2. Larger Developer and Community Reach
3. Increased Marketing
4. Inter App Integration
5. Reduced Cost of Development
6. Higher Success Ratio
7. Rich Development Environment

| Fact to test | Source |
|---|---|
| Checkbox: select the genuine reasons from a mixed list | p9 |
| Which reason relates to a *bigger pool of developers/community* | p9 |
| Which reason relates to *lower cost of building the app* | p9 |
| Which reason relates to *connecting with other apps* | p9 |
| Which reason relates to *a richer set of tools to build with* | p9 |

### Section D — The Android Story (Pages 10–16) — **5 questions**

| Fact to test | Source |
|---|---|
| Founder: **Andy Rubin**, Android Inc. | p12 |
| Founded in **Palo Alto, California, October 2003** | p13 |
| Co-founders: **Rich Miner, Nick Sears, Chris White** | p13 |
| Original goal: a "smarter mobile device more aware of its owner's location and preferences" | p13 |
| The company worked **secretively**, admitting only that it made software for mobile phones | p13 |

### Section E — Android beyond phones (Pages 17–29) — **11 questions**

| Fact to test | Source |
|---|---|
| Android TV is designed for **digital media players** | p20 |
| Android TV is a **replacement for Google TV**, built around content discovery + voice search | p20 |
| Android TV integrates Assistant, **Cast** and **Knowledge Graph** | p20 |
| Android Auto **mirrors** phone features onto the car's in-dash head unit | p22 |
| Checkbox: Android Auto's supported app categories (GPS/navigation, music, SMS, telephone, web search) | p22 |
| Android Auto supports **touchscreen and button-controlled** head units; hands-free voice is encouraged to reduce distraction | p22 |
| Wear OS was **previously called Android Wear** | p24 |
| Wear OS pairs with **Android 4.3+ or iOS 8.2+** (limited) | p24 |
| Checkbox: Wear OS connectivity (Bluetooth, Wi-Fi, 3G, LTE) | p24 |
| **Android Things** is the IoT platform — boards shown: Intel Edison, Raspberry Pi 3, NXP Pico | p26 |
| Android has been put into non-phone devices: washing machine, microwave (Touch Revolution, CES 2010), Samsung Galaxy Camera, Galaxy Tab | p25, p27–29 |

### Section F — Android Versions (Pages 30–52) — **23 questions**

**Overview (p31) — 3 questions**

| Fact to test |
|---|
| Over **2 billion** monthly active Android devices |
| Android is open source, written in Linux, launched by Google in **2008** |
| Every version after **1.5** has a code name chosen **alphabetically** |

**One question per Android version as the module presents them — 20 questions**

| Version | Code name | Signature fact tested |
|---|---|---|
| 1.0 | *(none)* | First release; only **35 apps**; had Gmail, Maps, Calendar, pull-down notifications, Android Market |
| 1.5 | Cupcake | **First version to get a code name**; animated transition effects |
| 1.6 | Donut | 800 × 480 screen support; voice search; handwriting gestures |
| 2.0/2.1 | Eclair | Media focus (flash, scene mode, white balance, digital zoom); **pinch-to-zoom** |
| 2.2 | Froyo | Install apps to **SD card**; USB tethering and Wi-Fi hotspot |
| 2.3 | Gingerbread | **NFC**; internet telephone and video chat |
| 3.0/3.1/3.2 | Honeycomb | Built **mainly for tablets**; multiple browser tabs, incognito mode |
| 4.0 | Ice Cream Sandwich | **Face Unlock**; multicore CPU support; virtual buttons |
| 4.1/4.2/4.3 | Jelly Bean | Smoother UI; world clock/stopwatch/timer; **native emoji** |
| 4.4 | KitKat | **"Ok, Google"**; printing framework; translucent UI; low-power sensors |
| 5.0/5.1 | Lollipop | Performance focus; **card-based concept**; screen capturing and sharing |
| 6.0 | Marshmallow | **Fingerprint reader support** and **USB-C** |
| 7.0 | Nougat | **Multi-Window**; Data Saver; Google Assistant |
| 8.0 | Oreo | **Picture-in-Picture**; notification channels; Autofill Framework |
| 9.0 | Pie | **Gesture/button navigation** replacing Back/Home/Overview; Digital Wellbeing |
| 10 | Quince Tart | **First to drop the dessert name**, known by number only; granular location permissions; dark theme |
| 11 | Red Velvet Cake | **One-time (single-use) permissions** for location, camera, microphone |
| 12 | Snow Cone | **Material You** — themes generated from your wallpaper colours |
| 14 | Upside Down Cake | Drag-and-drop text between apps; lock screen customisation; health data dashboard |
| 15 | Vanilla Ice Cream | *(see gap #2 below — no features stated in the module)* |

### Section G — Getting Started with Android Development (Pages 53–57) — **3 questions**

| Fact to test | Source |
|---|---|
| **Android Studio** is the IDE used (module shows Android Studio 3.0) | p55–56 |
| You need a **test device or a simulator** to run and test your app | p55, p57 |
| Why a simulator/emulator is useful when you have no physical device | p57 |

**Total: 2 + 6 + 5 + 5 + 11 + 23 + 3 = 55**

---

## 3. Gaps and problems I found in the source

These are things I *cannot* fairly write a question about, or that would produce a misleading
question. I want your decision on each rather than silently guessing.

| # | Issue | Pages | Why it matters |
|---|---|---|---|
| 1 | **Android 13 is missing entirely.** The deck jumps from Android 12 (p50) straight to Android 14 (p51). | 50→51 | Android 13 is Tiramisu. Either it was skipped in your lesson, or a slide is missing. I can leave it out (matching your module) or add it. |
| 2 | **Android 15 has no content** — only a heading and a mascot image, no feature list. | 52 | I can only ask "which is the newest version covered" — nothing about its features. |
| 3 | **Page 31 contradicts the rest of the deck.** It says Android has "about 8 major versions" and that "Oreo (8.0) is a current version", yet the deck runs to Android 15. | 31 | If I write a question from this slide, the "correct" answer is now false. I plan to quiz only the **stable** facts from p31 (2 billion devices, 2008, alphabetical code names) and skip the "8 versions / Oreo is current" claim. |
| 4 | **The alphabetical code-name rule breaks later.** Names are alphabetical up to Pie, then Android 10 onward used internal dessert names (Quince Tart, Red Velvet Cake, Snow Cone…) that Google stopped publicising. | 31, 48–52 | A question like "what comes after Pie alphabetically?" would be ambiguous. I'll only test that names *after 1.5* were chosen alphabetically, exactly as stated. |
| 5 | **Two names for one store.** p5 says "Android Market", p6 and most of the deck say "Google Play Store". | 5–6 | Same store, renamed in 2012. I'll test the rename as one fact rather than pretending they're different things. |
| 6 | **The downloads table is malformed.** The markdown has merged columns (e.g. `| 1 | TikTok * | 16M | Instagram * | 41M | TikTok | 49M |` — 7 cells under a 4-column header). | 6 | I can only trust the **Combined Total** column. I'll write one question from it (TikTok at 49M) and leave the per-store figures alone. |
| 7 | **Many slides are images only, with no text at all.** | 1, 7, 10, 14, 15, 16, 17, 18, 19, 21, 23, 27, 54 | There's nothing stated to quiz. p15 shows Steve Jobs with an early iPhone and p14 an early QWERTY Android phone, so I could write one context question about Android vs iPhone — but that would be my inference, not your module's text. Tell me if you want it. |
| 8 | **Some slides are headings only** — "Android-Powered Washing Machine" (p25), "Galaxy Tablet" (p27) has an image but no text. | 25, 27 | I'll cover them as one combined "Android runs on non-phone devices" question rather than inventing detail. |

---

## 4. What I need from you

Four decisions, then I'll write the quiz:

1. **How many questions?** 40 (core only) / **55 (recommended, full coverage)** / 70 (exhaustive)
2. **Android 13** — leave it out to match your module, or add Tiramisu?
3. **One quiz or several?** 55 in one file, or split into parts like SocPro?
4. **The "Most Downloaded Apps" table** — include it, or skip it as trivia?

---

## 5. After you approve

1. I'll write `WebDev_Mod1Raw.txt` in the exact `Question: / Option: / Explanation:` format.
2. Run it through `Convert-Quizzes.ps1` to produce `Quizizz_WebDev_Mod1.xlsx`.
3. Give you the conversion report, which flags any question with a missing answer or too many options.
4. You import the `.xlsx` into Wayground.
