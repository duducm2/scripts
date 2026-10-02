# -*- coding: utf-8 -*-
"""Add the five entertainment canons and their covers."""
import csv
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT / "mnemonics" / "data" / "entertainment.csv"
BFI = Path(
    r"C:\Users\eduev\.cursor\projects\c-Users-eduev-Meu-Drive-17-Projects-scripts\agent-tools\57728765-84fb-4a20-a22b-356820db80ff.txt"
)
UA = "MemoryPalaceEntertainment/1.0 (personal catalog; local use)"
FIELDS = ["id", "parent_id", "title", "url", "notes", "done", "sort_order", "image"]

BBC_TV = [
    "Small Axe (2020)",
    "This Is England '86, '88 and '90 (2010-2015)",
    "Call My Agent! (2015-2020)",
    "Happy Valley (2014-)",
    "The Shield (2002-2008)",
    "The Big Bang Theory (2007-2019)",
    "The Young Pope (2016)",
    "Dark (2017-2020)",
    "The Underground Railroad (2021)",
    "House of Cards (2013-2018)",
    "Avatar: The Last Airbender (2005-2008)",
    "The Good Place (2016-2020)",
    "Pose (2018-2021)",
    "Detectorists (2014-2017)",
    "Orange Is the New Black (2013-2019)",
    "Mare of Easttown (2021)",
    "RuPaul's Drag Race (2009-)",
    "Stranger Things (2016-)",
    "24 (2001-2010)",
    "Battlestar Galactica (2004-2009)",
    "Enlightened (2011-2013)",
    "Gilmore Girls (2000-2007)",
    "Planet Earth (2006)",
    "Utopia (2013-2014)",
    "Babylon Berlin (2017-)",
    "Rick and Morty (2013-)",
    "American Crime Story (2016-)",
    "The Killing (2007-2012)",
    "Mindhunter (2017-2019)",
    "House (2004-2012)",
    "O.J.: Made in America (2016)",
    "Big Little Lies (2017-2019)",
    "Insecure (2016-2021)",
    "Normal People (2020)",
    "Narcos (2015-2017)",
    "How I Met Your Mother (2005-2014)",
    "The Comeback (2005-2014)",
    "The OA (2016-2019)",
    "Dexter (2006-2013)",
    "It's Always Sunny in Philadelphia (2005-)",
    "Westworld (2016-)",
    "Show Me a Hero (2015)",
    "Treme (2010-2013)",
    "Louie (2010-2015)",
    "Luther (2010-2019)",
    "Catastrophe (2015-2019)",
    "Hannibal (2013-2015)",
    "Crazy Ex-Girlfriend (2015-2019)",
    "Steven Universe (2013-2020)",
    "The Queen's Gambit (2020)",
]

BIG_READ = [
    ("The Secret Garden", "Frances Hodgson Burnett"),
    ("Of Mice and Men", "John Steinbeck"),
    ("The Stand", "Stephen King"),
    ("Anna Karenina", "Leo Tolstoy"),
    ("A Suitable Boy", "Vikram Seth"),
    ("The BFG", "Roald Dahl"),
    ("Swallows and Amazons", "Arthur Ransome"),
    ("Black Beauty", "Anna Sewell"),
    ("Artemis Fowl", "Eoin Colfer"),
    ("Crime and Punishment", "Fyodor Dostoevsky"),
    ("Noughts and Crosses", "Malorie Blackman"),
    ("Memoirs of a Geisha", "Arthur Golden"),
    ("A Tale of Two Cities", "Charles Dickens"),
    ("The Thorn Birds", "Colleen McCullough"),
    ("Mort", "Terry Pratchett"),
    ("The Magic Faraway Tree", "Enid Blyton"),
    ("The Magus", "John Fowles"),
    ("Good Omens", "Terry Pratchett and Neil Gaiman"),
    ("Guards! Guards!", "Terry Pratchett"),
    ("Lord of the Flies", "William Golding"),
    ("Perfume", "Patrick Süskind"),
    ("The Ragged Trousered Philanthropists", "Robert Tressell"),
    ("Night Watch", "Terry Pratchett"),
    ("Matilda", "Roald Dahl"),
    ("Bridget Jones's Diary", "Helen Fielding"),
    ("The Secret History", "Donna Tartt"),
    ("The Woman in White", "Wilkie Collins"),
    ("Ulysses", "James Joyce"),
    ("Bleak House", "Charles Dickens"),
    ("Double Act", "Jacqueline Wilson"),
    ("The Twits", "Roald Dahl"),
    ("I Capture the Castle", "Dodie Smith"),
    ("Holes", "Louis Sachar"),
    ("Gormenghast", "Mervyn Peake"),
    ("The God of Small Things", "Arundhati Roy"),
    ("Vicky Angel", "Jacqueline Wilson"),
    ("Brave New World", "Aldous Huxley"),
    ("Cold Comfort Farm", "Stella Gibbons"),
    ("Magician", "Raymond E. Feist"),
    ("On the Road", "Jack Kerouac"),
    ("The Godfather", "Mario Puzo"),
    ("The Clan of the Cave Bear", "Jean M. Auel"),
    ("The Colour of Magic", "Terry Pratchett"),
    ("The Alchemist", "Paulo Coelho"),
    ("Katherine", "Anya Seton"),
    ("Kane and Abel", "Jeffrey Archer"),
    ("Love in the Time of Cholera", "Gabriel García Márquez"),
    ("Girls in Love", "Jacqueline Wilson"),
    ("The Princess Diaries", "Meg Cabot"),
    ("Midnight's Children", "Salman Rushdie"),
]

# Rolling Stone 500, 2020 edition, ranks 1-100. (album, artist, year)
RS_ALBUMS = [
    ("What's Going On", "Marvin Gaye", "1971"),
    ("Pet Sounds", "The Beach Boys", "1966"),
    ("Blue", "Joni Mitchell", "1971"),
    ("Songs in the Key of Life", "Stevie Wonder", "1976"),
    ("Abbey Road", "The Beatles", "1969"),
    ("Nevermind", "Nirvana", "1991"),
    ("Rumours", "Fleetwood Mac", "1977"),
    ("Purple Rain", "Prince and the Revolution", "1984"),
    ("Blood on the Tracks", "Bob Dylan", "1975"),
    ("The Miseducation of Lauryn Hill", "Lauryn Hill", "1998"),
    ("Revolver", "The Beatles", "1966"),
    ("Thriller", "Michael Jackson", "1982"),
    ("I Never Loved a Man the Way I Love You", "Aretha Franklin", "1967"),
    ("Exile on Main St.", "The Rolling Stones", "1972"),
    ("It Takes a Nation of Millions to Hold Us Back", "Public Enemy", "1988"),
    ("London Calling", "The Clash", "1979"),
    ("My Beautiful Dark Twisted Fantasy", "Kanye West", "2010"),
    ("Highway 61 Revisited", "Bob Dylan", "1965"),
    ("To Pimp a Butterfly", "Kendrick Lamar", "2015"),
    ("Kid A", "Radiohead", "2000"),
    ("Born to Run", "Bruce Springsteen", "1975"),
    ("Ready to Die", "The Notorious B.I.G.", "1994"),
    ("The Velvet Underground & Nico", "The Velvet Underground", "1967"),
    ("Sgt. Pepper's Lonely Hearts Club Band", "The Beatles", "1967"),
    ("Tapestry", "Carole King", "1971"),
    ("Horses", "Patti Smith", "1975"),
    ("Enter the Wu-Tang (36 Chambers)", "Wu-Tang Clan", "1993"),
    ("Voodoo", "D'Angelo", "2000"),
    ("The White Album", "The Beatles", "1968"),
    ("Are You Experienced", "The Jimi Hendrix Experience", "1967"),
    ("Kind of Blue", "Miles Davis", "1959"),
    ("Lemonade", "Beyoncé", "2016"),
    ("Back to Black", "Amy Winehouse", "2006"),
    ("Innervisions", "Stevie Wonder", "1973"),
    ("Rubber Soul", "The Beatles", "1965"),
    ("Off the Wall", "Michael Jackson", "1979"),
    ("The Chronic", "Dr. Dre", "1992"),
    ("Blonde on Blonde", "Bob Dylan", "1966"),
    ("Remain in Light", "Talking Heads", "1980"),
    ("The Rise and Fall of Ziggy Stardust and the Spiders from Mars", "David Bowie", "1972"),
    ("Let It Bleed", "The Rolling Stones", "1969"),
    ("OK Computer", "Radiohead", "1997"),
    ("The Low End Theory", "A Tribe Called Quest", "1991"),
    ("Illmatic", "Nas", "1994"),
    ("Sign o' the Times", "Prince", "1987"),
    ("Graceland", "Paul Simon", "1986"),
    ("Ramones", "Ramones", "1976"),
    ("Legend", "Bob Marley and the Wailers", "1984"),
    ("Aquemini", "Outkast", "1998"),
    ("The Blueprint", "Jay-Z", "2001"),
    ("The Great Twenty-Eight", "Chuck Berry", "1982"),
    ("Station to Station", "David Bowie", "1976"),
    ("Electric Ladyland", "The Jimi Hendrix Experience", "1968"),
    ("Star Time", "James Brown", "1991"),
    ("The Dark Side of the Moon", "Pink Floyd", "1973"),
    ("Exile in Guyville", "Liz Phair", "1993"),
    ("The Band", "The Band", "1969"),
    ("Led Zeppelin IV", "Led Zeppelin", "1971"),
    ("Talking Book", "Stevie Wonder", "1972"),
    ("Astral Weeks", "Van Morrison", "1968"),
    ("Paid in Full", "Eric B. & Rakim", "1987"),
    ("Appetite for Destruction", "Guns N' Roses", "1987"),
    ("Aja", "Steely Dan", "1977"),
    ("Stankonia", "Outkast", "2000"),
    ("Live at the Apollo", "James Brown", "1963"),
    ("A Love Supreme", "John Coltrane", "1965"),
    ("Reasonable Doubt", "Jay-Z", "1996"),
    ("Hounds of Love", "Kate Bush", "1985"),
    ("Jagged Little Pill", "Alanis Morissette", "1995"),
    ("Straight Outta Compton", "N.W.A", "1988"),
    ("Exodus", "Bob Marley and the Wailers", "1977"),
    ("Harvest", "Neil Young", "1972"),
    ("Loveless", "My Bloody Valentine", "1991"),
    ("The College Dropout", "Kanye West", "2004"),
    ("Lady Soul", "Aretha Franklin", "1968"),
    ("Super Fly", "Curtis Mayfield", "1972"),
    ("Who's Next", "The Who", "1971"),
    ("The Sun Sessions", "Elvis Presley", "1976"),
    ("Blonde", "Frank Ocean", "2016"),
    ("Never Mind the Bollocks, Here's the Sex Pistols", "Sex Pistols", "1977"),
    ("Beyoncé", "Beyoncé", "2013"),
    ("There's a Riot Goin' On", "Sly and the Family Stone", "1971"),
    ("Dusty in Memphis", "Dusty Springfield", "1969"),
    ("Back in Black", "AC/DC", "1980"),
    ("Plastic Ono Band", "John Lennon", "1970"),
    ("The Doors", "The Doors", "1967"),
    ("Bitches Brew", "Miles Davis", "1970"),
    ("Hunky Dory", "David Bowie", "1971"),
    ("Baduizm", "Erykah Badu", "1997"),
    ("After the Gold Rush", "Neil Young", "1970"),
    ("Darkness on the Edge of Town", "Bruce Springsteen", "1978"),
    ("Axis: Bold as Love", "The Jimi Hendrix Experience", "1967"),
    ("Supa Dupa Fly", "Missy Elliott", "1997"),
    ("Fun House", "The Stooges", "1970"),
    ("Take Care", "Drake", "2011"),
    ("Automatic for the People", "R.E.M.", "1992"),
    ("Master of Puppets", "Metallica", "1986"),
    ("Car Wheels on a Gravel Road", "Lucinda Williams", "1998"),
    ("Red", "Taylor Swift", "2012"),
    ("Music from Big Pink", "The Band", "1968"),
]

RS_TV = [
    "The Sopranos (1999-2007)",
    "The Simpsons (1989-)",
    "Breaking Bad (2008-2013)",
    "The Wire (2002-2008)",
    "Fleabag (2016-2019)",
    "Seinfeld (1989-1998)",
    "Mad Men (2007-2015)",
    "Cheers (1982-1993)",
    "Atlanta (2016-)",
    "The Mary Tyler Moore Show (1970-1977)",
    "Succession (2018-)",
    "The Twilight Zone (1959-1964)",
    "Veep (2012-2019)",
    "The Americans (2013-2018)",
    "The Larry Sanders Show (1992-1998)",
    "Twin Peaks (1990-1991)",
    "The Leftovers (2014-2017)",
    "Saturday Night Live (1975-)",
    "I May Destroy You (2020)",
    "30 Rock (2006-2013)",
    "All in the Family (1971-1979)",
    "Star Trek (1966-1969)",
    "Watchmen (2019)",
    "Freaks and Geeks (1999-2000)",
    "M*A*S*H (1972-1983)",
    "Sesame Street (1969-)",
    "Deadwood (2004-2006)",
    "Friday Night Lights (2006-2011)",
    "Roots (1977)",
    "Parks and Recreation (2009-2015)",
    "Game of Thrones (2011-2019)",
    "Better Call Saul (2015-2022)",
    "Monty Python's Flying Circus (1969-1974)",
    "The Office (US) (2005-2013)",
    "Lost (2004-2010)",
    "I Love Lucy (1951-1957)",
    "Arrested Development (2003-2019)",
    "Hill Street Blues (1981-1987)",
    "Curb Your Enthusiasm (2000-)",
    "The Good Place (2016-2020)",
    "BoJack Horseman (2014-2020)",
    "Battlestar Galactica (2004-2009)",
    "Insecure (2016-2021)",
    "Late Night with David Letterman (1982-1993)",
    "Columbo (1971-1978)",
    "The West Wing (1999-2006)",
    "My So-Called Life (1994-1995)",
    "The Shield (2002-2008)",
    "Friends (1994-2004)",
    "Jeopardy! (1984-)",
    "The X-Files (1993-2018)",
    "Barry (2018-2023)",
    "The Office (UK) (2001-2003)",
    "ER (1994-2009)",
    "Halt and Catch Fire (2014-2017)",
    "Community (2009-2015)",
    "Russian Doll (2019-)",
    "Six Feet Under (2001-2005)",
    "Key & Peele (2012-2015)",
    "Taxi (1978-1983)",
    "The Underground Railroad (2021)",
    "The Dick Van Dyke Show (1961-1966)",
    "South Park (1997-)",
    "The Golden Girls (1985-1992)",
    "Girls (2012-2017)",
    "The Daily Show with Jon Stewart (1999-2015)",
    "NYPD Blue (1993-2005)",
    "Fawlty Towers (1975-1979)",
    "Chappelle's Show (2003-2006)",
    "SCTV (1976-1984)",
    "Better Things (2016-2022)",
    "Good Times (1974-1979)",
    "Buffy the Vampire Slayer (1997-2003)",
    "The Honeymooners (1955-1956)",
    "Frasier (1993-2004)",
    "Justified (2010-2015)",
    "The Jeffersons (1975-1985)",
    "Sex and the City (1998-2004)",
    "Mr. Show with Bob and David (1995-1998)",
    "Band of Brothers (2001)",
    "It's Always Sunny in Philadelphia (2005-)",
    "Party Down (2009-)",
    "I'm Alan Partridge (1997-2002)",
    "Fargo (2014-)",
    "Orange Is the New Black (2013-2019)",
    "The Bob Newhart Show (1972-1978)",
    "The Kids in the Hall (1988-1995)",
    "The Crown (2016-)",
    "The Carol Burnett Show (1967-1978)",
    "The Wonder Years (1988-1993)",
    "The Tonight Show Starring Johnny Carson (1962-1992)",
    "The Muppet Show (1976-1981)",
    "The Rockford Files (1974-1980)",
    "NewsRadio (1995-1999)",
    "Squid Game (2021-)",
    "Rick and Morty (2013-)",
    "The Odd Couple (1970-1975)",
    "The Good Fight (2017-2022)",
    "Oz (1997-2003)",
    "What We Do in the Shadows (2019-)",
]

# Don Quixote first, then the Wikipedia list by author, without a second Don Quixote.
WORLD = [
    ("Don Quixote", "Miguel de Cervantes"),
    ("Things Fall Apart", "Chinua Achebe"),
    ("Fairy Tales", "Hans Christian Andersen"),
    ("The Divine Comedy", "Dante Alighieri"),
    ("The Epic of Gilgamesh", "Unknown"),
    ("The Book of Job", "Unknown"),
    ("One Thousand and One Nights", "Various"),
    ("Njál's Saga", "Unknown"),
    ("Pride and Prejudice", "Jane Austen"),
    ("Père Goriot", "Honoré de Balzac"),
    ("Molloy, Malone Dies, The Unnamable", "Samuel Beckett"),
    ("The Decameron", "Giovanni Boccaccio"),
    ("Ficciones", "Jorge Luis Borges"),
    ("Wuthering Heights", "Emily Brontë"),
    ("The Stranger", "Albert Camus"),
    ("Poems", "Paul Celan"),
    ("Journey to the End of the Night", "Louis-Ferdinand Céline"),
    ("The Canterbury Tales", "Geoffrey Chaucer"),
    ("Selected Stories", "Anton Chekhov"),
    ("Nostromo", "Joseph Conrad"),
    ("Great Expectations", "Charles Dickens"),
    ("Jacques the Fatalist", "Denis Diderot"),
    ("Berlin Alexanderplatz", "Alfred Döblin"),
    ("Crime and Punishment", "Fyodor Dostoevsky"),
    ("The Idiot", "Fyodor Dostoevsky"),
    ("Demons", "Fyodor Dostoevsky"),
    ("The Brothers Karamazov", "Fyodor Dostoevsky"),
    ("Middlemarch", "George Eliot"),
    ("Invisible Man", "Ralph Ellison"),
    ("Medea", "Euripides"),
    ("Absalom, Absalom!", "William Faulkner"),
    ("The Sound and the Fury", "William Faulkner"),
    ("Madame Bovary", "Gustave Flaubert"),
    ("Sentimental Education", "Gustave Flaubert"),
    ("Gypsy Ballads", "Federico García Lorca"),
    ("One Hundred Years of Solitude", "Gabriel García Márquez"),
    ("Love in the Time of Cholera", "Gabriel García Márquez"),
    ("Faust", "Johann Wolfgang von Goethe"),
    ("Dead Souls", "Nikolai Gogol"),
    ("The Tin Drum", "Günter Grass"),
    ("The Devil to Pay in the Backlands", "João Guimarães Rosa"),
    ("Hunger", "Knut Hamsun"),
    ("The Old Man and the Sea", "Ernest Hemingway"),
    ("The Iliad", "Homer"),
    ("The Odyssey", "Homer"),
    ("A Doll's House", "Henrik Ibsen"),
    ("Ulysses", "James Joyce"),
    ("Selected Stories", "Franz Kafka"),
    ("The Trial", "Franz Kafka"),
    ("The Castle", "Franz Kafka"),
    ("Shakuntala", "Kālidāsa"),
    ("The Sound of the Mountain", "Yasunari Kawabata"),
    ("Zorba the Greek", "Nikos Kazantzakis"),
    ("Sons and Lovers", "D. H. Lawrence"),
    ("Independent People", "Halldór Laxness"),
    ("Complete Poems", "Giacomo Leopardi"),
    ("The Golden Notebook", "Doris Lessing"),
    ("Pippi Longstocking", "Astrid Lindgren"),
    ("Diary of a Madman and Other Stories", "Lu Xun"),
    ("Children of Gebelawi", "Naguib Mahfouz"),
    ("Buddenbrooks", "Thomas Mann"),
    ("The Magic Mountain", "Thomas Mann"),
    ("Moby-Dick", "Herman Melville"),
    ("Essays", "Michel de Montaigne"),
    ("History", "Elsa Morante"),
    ("Beloved", "Toni Morrison"),
    ("The Tale of Genji", "Murasaki Shikibu"),
    ("The Man Without Qualities", "Robert Musil"),
    ("Lolita", "Vladimir Nabokov"),
    ("Nineteen Eighty-Four", "George Orwell"),
    ("Metamorphoses", "Ovid"),
    ("The Book of Disquiet", "Fernando Pessoa"),
    ("Tales", "Edgar Allan Poe"),
    ("In Search of Lost Time", "Marcel Proust"),
    ("Gargantua and Pantagruel", "François Rabelais"),
    ("Pedro Páramo", "Juan Rulfo"),
    ("Masnavi", "Rumi"),
    ("Midnight's Children", "Salman Rushdie"),
    ("Bostan", "Saadi"),
    ("Season of Migration to the North", "Tayeb Salih"),
    ("Blindness", "José Saramago"),
    ("Hamlet", "William Shakespeare"),
    ("King Lear", "William Shakespeare"),
    ("Othello", "William Shakespeare"),
    ("Oedipus the King", "Sophocles"),
    ("The Red and the Black", "Stendhal"),
    ("Tristram Shandy", "Laurence Sterne"),
    ("Confessions of Zeno", "Italo Svevo"),
    ("Gulliver's Travels", "Jonathan Swift"),
    ("War and Peace", "Leo Tolstoy"),
    ("Anna Karenina", "Leo Tolstoy"),
    ("The Death of Ivan Ilyich", "Leo Tolstoy"),
    ("Adventures of Huckleberry Finn", "Mark Twain"),
    ("The Ramayana", "Valmiki"),
    ("The Aeneid", "Virgil"),
    ("The Mahabharata", "Vyasa"),
    ("Leaves of Grass", "Walt Whitman"),
    ("Mrs Dalloway", "Virginia Woolf"),
    ("To the Lighthouse", "Virginia Woolf"),
    ("Memoirs of Hadrian", "Marguerite Yourcenar"),
]

FILM_ALIAS = {
    "Jeanne Dielman, 23 Quai du Commerce, 1080 Bruxelles": "Jeanne Dielman, 23 Quai du Commerce, 1080 Bruxelles",
    "Sunrise A Song of Two Humans": "Sunrise: A Song of Two Humans",
    "La Règle du jeu": "The Rules of the Game",
    "À bout de souffle": "Breathless",
    "8½": "8½",
}


def norm(text):
    text = text.lower().replace("’", "'").replace("–", "-").replace("—", "-")
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def core(title):
    text = title.split("—")[0].strip()
    text = re.sub(r"\s*\([^)]*(?:19|20)\d{2}[^)]*\)\s*$", "", text)
    return norm(text)


def book_title(title, author):
    return f"{title} — {author}"


def directors():
    text = BFI.read_text(encoding="utf-8")
    blocks = re.split(r"\n# ", text)
    films = []
    for block in blocks:
        title = block.split("\n", 1)[0].strip().lstrip("# ").strip()
        if not title or "Greatest Films" in title or title.startswith("Stream"):
            continue
        years = re.findall(r"(?:1[89]\d{2}|20\d{2})", block)
        glued = re.findall(r"=\d{1,3}((?:1[89]|20)\d{2})", block)
        years = years + glued
        if not years:
            continue
        title = FILM_ALIAS.get(title, title)
        films.append(f"{title} ({years[-1]})")
    films = list(reversed(films))
    return films[:100]


def already(rows, parent_id, title):
    key = norm(title)
    return any(row["parent_id"] == parent_id and norm(row["title"]) == key for row in rows)


def add_row(rows, next_n, parent_id, title, sort_order, url=""):
    row = {
        "id": f"ENT_{next_n:04d}",
        "parent_id": parent_id,
        "title": title,
        "url": url,
        "notes": "",
        "done": "0",
        "sort_order": str(sort_order),
        "image": "",
    }
    rows.append(row)
    return next_n + 1, row


def get_json(url):
    delay = 2.0
    for attempt in range(5):
        req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.load(resp)
        except urllib.error.HTTPError as err:
            if err.code in (429, 503) and attempt + 1 < 5:
                time.sleep(delay)
                delay = min(delay * 2, 20)
                continue
            raise
        except (urllib.error.URLError, TimeoutError):
            if attempt + 1 < 5:
                time.sleep(delay)
                continue
            raise
    return {}


def tokens(text):
    stop = {"the", "a", "an", "of", "and", "vol"}
    return {w for w in norm(text).split() if w not in stop and len(w) >= 3}


def wiki_image(query, want):
    url = (
        "https://en.wikipedia.org/w/api.php?action=query&format=json"
        "&generator=search&gsrlimit=4&gsrnamespace=0&prop=pageimages"
        "&piprop=thumbnail&pithumbsize=500&pilicense=any&gsrsearch="
        + urllib.parse.quote(query)
    )
    data = get_json(url)
    pages = list(((data.get("query") or {}).get("pages") or {}).values())
    pages.sort(key=lambda p: p.get("index", 99))
    for page in pages:
        thumb = ((page.get("thumbnail") or {}).get("source") or "").split("?")[0]
        if not thumb or ".svg" in thumb.lower():
            continue
        if want & tokens(page.get("title") or ""):
            return thumb
    return ""


def itunes_cover(album, artist):
    url = (
        "https://itunes.apple.com/search?term="
        + urllib.parse.quote(f"{album} {artist}")
        + "&entity=album&limit=6&country=us"
    )
    data = get_json(url)
    want_album = norm(album)
    want_artist = tokens(artist)
    best = None
    best_score = 0
    for item in data.get("results") or []:
        name = norm(item.get("collectionName") or "")
        if any(word in name for word in ("karaoke", "tribute")):
            continue
        score = 0
        if name == want_album or name.startswith(want_album + " "):
            score += 10
        elif want_album and want_album in name:
            score += 4
        if want_artist & tokens(item.get("artistName") or ""):
            score += 6
        if score > best_score:
            best_score = score
            best = item
    if not best or best_score < 10:
        return ""
    return (best.get("artworkUrl100") or "").replace("100x100bb", "600x600bb").split("?")[0]


def tvmaze_cover(title):
    show = re.sub(r"\s*\([^)]*\)\s*$", "", title).strip()
    url = "https://api.tvmaze.com/search/shows?q=" + urllib.parse.quote(show)
    data = get_json(url)
    want = norm(show)
    for item in data[:5]:
        info = item.get("show") or {}
        if norm(info.get("name") or "") != want and not norm(info.get("name") or "").startswith(want):
            continue
        image = ((info.get("image") or {}).get("original") or "").split("?")[0]
        if image:
            return image
    return ""


def openlibrary_cover(title, author):
    url = (
        "https://openlibrary.org/search.json?title="
        + urllib.parse.quote(title)
        + "&author="
        + urllib.parse.quote(author)
        + "&limit=3"
    )
    data = get_json(url)
    want = tokens(title)
    for doc in data.get("docs") or []:
        if want and not (want & tokens(doc.get("title") or "")):
            continue
        cover = doc.get("cover_i")
        if cover:
            return f"https://covers.openlibrary.org/b/id/{cover}-L.jpg"
    return ""


def fetch_cover(row):
    title = row["title"]
    parent = row["parent_id"]
    kind = row.get("_kind")
    if kind == "album":
        album, artist = [part.strip() for part in title.split("—", 1)]
        album = re.sub(r"\s*\([^)]*\)\s*$", "", album).strip()
        artist = re.sub(r"\s*\([^)]*\)\s*$", "", artist).strip()
        image = itunes_cover(album, artist)
        if image:
            return image
        time.sleep(0.3)
        return wiki_image(f'"{album}" {artist} album', tokens(album) or tokens(artist))
    if kind == "tv":
        image = tvmaze_cover(title)
        if image:
            return image
        time.sleep(0.3)
        show = re.sub(r"\s*\([^)]*\)\s*$", "", title)
        return wiki_image(f'"{show}" television series', tokens(show))
    if kind == "book":
        name, author = [part.strip() for part in title.split("—", 1)]
        image = openlibrary_cover(name, author)
        if image:
            return image
        time.sleep(0.3)
        return wiki_image(f'"{name}" {author} book', tokens(name))
    name = re.sub(r"\s*\([^)]*\)\s*$", "", title)
    year = re.search(r"\((\d{4})", title)
    year = year.group(1) if year else ""
    return wiki_image(f'"{name}" {year} film', tokens(name))


def reuse(row, library):
    key = core(row["title"])
    if len(key) < 6:
        return ""
    return library.get(key, "")


def main():
    with CSV_PATH.open(encoding="utf-8", newline="") as handle:
        rows = list(csv.DictReader(handle))
    for row in rows:
        row.setdefault("image", "")
    if any(r["title"] == "Sight and Sound directors" for r in rows):
        raise SystemExit("canons already added")
    next_n = max(int(r["id"].split("_")[1]) for r in rows) + 1
    new_rows = []

    kids = [r for r in rows if r["parent_id"] == "ENT_0017"]
    sort = max(int(r["sort_order"] or 0) for r in kids)
    for title in BBC_TV:
        if already(rows, "ENT_0017", title):
            continue
        sort += 1
        next_n, row = add_row(rows, next_n, "ENT_0017", title, sort)
        row["_kind"] = "tv"
        new_rows.append(row)

    kids = [r for r in rows if r["parent_id"] == "ENT_0033"]
    sort = max(int(r["sort_order"] or 0) for r in kids)
    for title, author in BIG_READ:
        label = book_title(title, author)
        if already(rows, "ENT_0033", label):
            continue
        sort += 1
        next_n, row = add_row(rows, next_n, "ENT_0033", label, sort)
        row["_kind"] = "book"
        new_rows.append(row)

    for row in rows:
        if row["parent_id"] == "ENT_0001" and str(row["sort_order"]).isdigit() and int(row["sort_order"]) >= 3:
            row["sort_order"] = str(int(row["sort_order"]) + 1)
    next_n, heading = add_row(
        rows,
        next_n,
        "ENT_0001",
        "Sight and Sound directors",
        3,
        "https://www.bfi.org.uk/sight-and-sound/directors-100-greatest-films-all-time",
    )
    film_titles = directors()
    if len(film_titles) != 100:
        raise SystemExit(f"directors count {len(film_titles)}")
    for rank, title in enumerate(film_titles, 1):
        next_n, row = add_row(rows, next_n, heading["id"], title, rank)
        row["_kind"] = "film"
        new_rows.append(row)

    for row in rows:
        if row["parent_id"] == "ENT_0034" and str(row["sort_order"]).isdigit() and int(row["sort_order"]) >= 2:
            row["sort_order"] = str(int(row["sort_order"]) + 1)
    next_n, heading = add_row(
        rows,
        next_n,
        "ENT_0034",
        "Rolling Stone albums",
        2,
        "https://www.rollingstone.com/music/music-lists/best-albums-of-all-time-1062063/",
    )
    for rank, (album, artist, year) in enumerate(RS_ALBUMS, 1):
        next_n, row = add_row(rows, next_n, heading["id"], f"{album} — {artist} ({year})", rank)
        row["_kind"] = "album"
        new_rows.append(row)

    next_n, heading = add_row(
        rows,
        next_n,
        "ENT_0017",
        "Rolling Stone",
        0,
        "https://www.rollingstone.com/tv-movies/tv-movie-lists/best-tv-shows-of-all-time-1234598313/",
    )
    for rank, title in enumerate(RS_TV, 1):
        next_n, row = add_row(rows, next_n, heading["id"], title, rank)
        row["_kind"] = "tv"
        new_rows.append(row)

    next_n, heading = add_row(
        rows,
        next_n,
        "ENT_0033",
        "World Library",
        0,
        "https://en.wikipedia.org/wiki/Bokklubben_World_Library",
    )
    for rank, (title, author) in enumerate(WORLD, 1):
        next_n, row = add_row(rows, next_n, heading["id"], book_title(title, author), rank)
        row["_kind"] = "book"
        new_rows.append(row)

    library = {}
    for row in rows:
        if row.get("image") and row not in new_rows:
            library.setdefault(core(row["title"]), row["image"])
    reused = 0
    for row in new_rows:
        image = reuse(row, library)
        if image:
            row["image"] = image
            reused += 1
    print("rows", len(new_rows), "reused", reused, "fetch", len(new_rows) - reused, flush=True)

    def save():
        with CSV_PATH.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=FIELDS, lineterminator="\n", extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)

    save()
    pending = [r for r in new_rows if not r.get("image")]
    for i, row in enumerate(pending, 1):
        try:
            row["image"] = fetch_cover(row)
        except Exception as err:
            print("ERR", row["title"], err, flush=True)
            row["image"] = ""
        if not row["image"]:
            print("MISS", row["title"], flush=True)
        time.sleep(0.35)
        if i % 20 == 0:
            save()
            print(f"covers {i}/{len(pending)}", flush=True)
    save()
    filled = sum(1 for r in new_rows if r.get("image"))
    print("DONE", filled, "/", len(new_rows), flush=True)


if __name__ == "__main__":
    main()
