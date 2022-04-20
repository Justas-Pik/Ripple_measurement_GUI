# Ripple_measurement_GUI
GUI for ripple measurement system.
Maitinimo šaltinių testavimo sistema

Aprašymas vartojojui:
Naudojantis tinklo šakotuvu (ang. Switch) bei LAN sąsajomis prijungiame osciloskopą, elektroninę apkrovą ir tinklo sietuvą prie kompiuterio. Įjungiame tinklo šakotuvą, sietuvą ir testavimo stendą į elektros tinklą. Sujungiame tinklo sietuvą su testavimo stendu, trimis laideliais, skirtais komutuoti maitinimo šaltinių kanalus. Prie testavimo stendo prijungiame osciloskopą ir elektroninę apkrovą, taip pat ne daugiau kaip 4 maitinimo šaltinius. Kompiuteryje įjungiame vartotojo sąsaja. Suvedame ir nustatome įrangos IP adresus bei tinklo sietuvo slaptažodį, jei duomenys nebuvo keičiami įvestos pradinės reikšmės turėtų tikti. Pasirenkami testavimo stendo kanalai bei matavimų tipai. Įvedama kiekvieno kanalo apkrovos srovė. Matavimai pradedami mygtuku „Start“.  Atlikus pirmąjį matavimą reikia pasirinkti kuriame aplanke išsaugoti dokumentą, nepasirinkus dokumentas išsaugomas tik tame aplanke, kuriame yra vartotojo sąsajos failas, pasirinkus papildomą aplanką, duomenis išsaugomi ir ten ir ten, yra galimybė pasirinkti papildomą aplanką prieš matavimą, tuomet matavimo metu nebus klausiama, kur išsaugoti dokumentą. Atlikus matavimus nuo prietaisų atsijungiama mygtuku „Close“. Programa uždaroma mygtuku „X“ arba „Quit“.

Pastabos: 
Įjungus vartotojo sąsają, tik prisijungus prie įrangos leidžiama pasirinkti matavimus, bei kanalus. Įvedus srovės reikšmę, tačiau norint ją pakeisti, reikia paspausti mygtuką „Cancel“, tas pats galioja mygtukui „Submit“, kai keičiami dinaminės apkrovos parametrai. Yra galimybė pakeisti šiuos dinaminius parametrus: A_lvl, A_width, B_width, Ris_slew, Fal_slew.
Dinaminiame matavime B lygis atitinka įvestą prie matuojamo kanalo lygį.
Keičiant dinaminus parametrus nuspaudžiamas mygtukas „Change“, pakeičiamos vertės ir patvirtinama mygtuku „Submit“. Pakeitus dinaminius parametrus ir paspaudus mygtuką „Change“ surašomos pradinės reikšmės.

Privalumai:
Galimybė pasirinkti kanalus ir matavimus.
Srovės įvedimas kiekvienam kanalui.
Galimybė dokumentą įrašyti į iš anksto pasirinktą aplanką, jei nepasirinktas atlikus pirmą matavimą paklausiama, į kurį aplanką įrašyti.
Atliekant No Load, Full Load ir Dynamic Load matavimus  išvedama papildoma lentelė su pertvarkytais duomenimis.
Galimybė keisti dinaminius parametrus.
Atlikus matavimus išvedamos žinutės, taip pat, jei neįvesta srovė ir norima atlikti pilnos ar dinaminės apkrovos matavimus, suvedus elektroninės apkrovos rėžių neatitinkančius dinaminus parametrus ar raides vietoj adreso.

Trūkumai:
Failas automatiškai išsaugomas tame aplanke, kuriame yra .exe failas, nurodant vietą išsaugomas ir nurodytame aplanke ir .exe failo aplanke. Jei atliekant "No Load" matavimai su keliais kanalais nenurodant papildomo kelio (t.y iššokusiame lange paspaudus „Cancel“), kiekvienas matavimas išskaidomas į atskirą dokumentą ir patalpinamas aplanke, kuriame yra vartotojo sąsaja. Pasirinkus "No Load" ir "Full Load"  matavimas, ir nenurodant kelio, matavimas po "Full Load" nutraukiamas, srovė lieka įjungta, dokumentas patalpinamas vartotojo sąsajos aplanke.
Srovių įvedimo laukeliai (Current CH1,2,3,4) nėra apsaugoti nuo raidžių įvedimo, jas įvedus aprkovoje nustatoma 0 A srovė.
Žinutės neišvedamos realiu laiku.
Uždarant aplikaciją neuždaromi ryšiai su prietaisais. Atsijungiama nuspaudus „Close“ mygtuką vartotojo sąsajoje, tik tuomet galima uždaryti programą spaudžiant viršutiniame kampe esantį „X“ arba „Quit“ mygtukus.
Long term užkomentuotas, atkomentavus ir paleidus komandą vyktų kalibravimas tada 1h laukimas ir tik tada matavimas ir duomenų įrašymas.
Nutraukus ryšius su prietaisais nėra suvaldomos klaidos.
