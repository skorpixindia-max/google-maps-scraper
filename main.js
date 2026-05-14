import { Actor } from 'apify';
import { PlaywrightCrawler, sleep, log } from 'crawlee';
import { existsSync, readFileSync, writeFileSync, mkdirSync } from 'fs';
import { join } from 'path';
import * as readline from 'readline';
import ExcelJS from 'exceljs';

// ════════════════════════════════════════════════════════════════
//   USER AGENT & VIEWPORT ROTATION
// ════════════════════════════════════════════════════════════════
const USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:124.0) Gecko/20100101 Firefox/124.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 Edg/124.0.0.0',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
];
const VIEWPORTS = [
    { width:1366, height:768 }, { width:1440, height:900 },
    { width:1536, height:864 }, { width:1280, height:720 },
];
let uaIdx = 0, vpIdx = 0;
const nextUA       = () => USER_AGENTS[uaIdx++ % USER_AGENTS.length];
const nextViewport = () => VIEWPORTS[vpIdx++ % VIEWPORTS.length];
const randDelay    = (min=400,max=900) => Math.floor(Math.random()*(max-min)+min);

// ════════════════════════════════════════════════════════════════
//   BIG CHAINS TO SKIP
// ════════════════════════════════════════════════════════════════
const BIG_CHAINS = [
    'apollo','manipal','fortis','max hospital','columbia asia','narayana',
    'medanta','aiims','kims','care hospital','yashoda','continental',
    'star hospital','rainbow','mcdonalds','kfc','dominos','pizza hut',
    'subway','starbucks','walmart','target','costco','cvs pharmacy',
];

// ════════════════════════════════════════════════════════════════
//   SELECTORS
// ════════════════════════════════════════════════════════════════
const SEL = {
    phone    : 'button[data-item-id^="phone"] .Io6YTe',
    title    : 'h1.DUwDvf, h1[class*="fontHeadlineLarge"]',
    category : 'button[jsaction*="category"]',
    website  : 'a[data-item-id="authority"]',
    feed     : 'div[role="feed"]',
};

// ════════════════════════════════════════════════════════════════
//   GEO-GRID: known area center coordinates [lat, lng]
// ════════════════════════════════════════════════════════════════
const AREA_COORDS = {
    // Chennai areas
    't nagar':[13.0418,80.2341],'anna nagar':[13.0891,80.2096],'velachery':[12.9815,80.2180],
    'adyar':[13.0012,80.2565],'nungambakkam':[13.0569,80.2425],'tambaram':[12.9249,80.1000],
    'porur':[13.0358,80.1567],'chromepet':[12.9516,80.1462],'perambur':[13.1166,80.2383],
    'mogappair':[13.0827,80.1742],'kodambakkam':[13.0516,80.2253],'mylapore':[13.0368,80.2676],
    'royapettah':[13.0523,80.2601],'egmore':[13.0782,80.2637],'ambattur':[13.1143,80.1548],
    // Bangalore areas
    'koramangala':[12.9352,77.6245],'indiranagar':[12.9784,77.6408],'whitefield':[12.9698,77.7499],
    'hsr layout':[12.9116,77.6389],'jp nagar':[12.9077,77.5838],'jayanagar':[12.9254,77.5938],
    'electronic city':[12.8452,77.6602],'marathahalli':[12.9591,77.6974],'rajajinagar':[12.9908,77.5530],
    'malleshwaram':[13.0031,77.5705],'yelahanka':[13.1007,77.5963],'bannerghatta road':[12.8933,77.5975],
    'sarjapur road':[12.9122,77.6841],'hebbal':[13.0450,77.5950],'rt nagar':[13.0218,77.5972],
    // Hyderabad areas
    'banjara hills':[17.4126,78.4480],'jubilee hills':[17.4312,78.4071],'madhapur':[17.4400,78.3912],
    'kondapur':[17.4605,78.3575],'kukatpally':[17.4947,78.3996],'ameerpet':[17.4374,78.4482],
    'gachibowli':[17.4401,78.3489],'secunderabad':[17.4399,78.4983],'lb nagar':[17.3490,78.5497],
    'dilsukhnagar':[17.3688,78.5282],'uppal':[17.4051,78.5590],'miyapur':[17.4966,78.3547],
    // Mumbai areas
    'andheri':[19.1136,72.8697],'bandra':[19.0596,72.8295],'dadar':[19.0178,72.8478],
    'borivali':[19.2307,72.8567],'malad':[19.1874,72.8484],'powai':[19.1176,72.9060],
    'juhu':[19.1075,72.8263],'thane':[19.2183,72.9781],'kurla':[19.0726,72.8793],
    'goregaon':[19.1663,72.8526],'mulund':[19.1726,72.9572],'chembur':[19.0626,72.8991],
    // Pune areas
    'koregaon park':[18.5362,73.8938],'kothrud':[18.5074,73.8077],'wakad':[18.5975,73.7617],
    'baner':[18.5590,73.7868],'hadapsar':[18.5018,73.9260],'viman nagar':[18.5679,73.9143],
    // Delhi areas
    'connaught place':[28.6315,77.2167],'lajpat nagar':[28.5677,77.2434],'karol bagh':[28.6514,77.1907],
    'dwarka':[28.5921,77.0460],'rohini':[28.7041,77.1025],'saket':[28.5244,77.2090],
    'vasant kunj':[28.5200,77.1588],'nehru place':[28.5491,77.2529],
    // USA
    'manhattan':[40.7831,-73.9712],'brooklyn':[40.6782,-73.9442],'queens':[40.7282,-73.7949],
    'downtown los angeles':[34.0407,-118.2468],'hollywood':[34.0928,-118.3287],'santa monica':[34.0195,-118.4912],
    'downtown chicago':[41.8827,-87.6233],'lincoln park':[41.9214,-87.6513],
    'downtown miami':[25.7617,-80.1918],'south beach':[25.7826,-80.1340],
    'downtown houston':[29.7604,-95.3698],'midtown manhattan':[40.7549,-73.9840],
    // UK
    'central london':[51.5074,-0.1278],'shoreditch':[51.5227,-0.0798],
    'canary wharf':[51.5054,-0.0235],'city of london':[51.5155,-0.0922],
    'central manchester':[53.4808,-2.2426],'central birmingham':[52.4862,-1.8904],
    // UAE
    'dubai marina':[25.0805,55.1403],'downtown dubai':[25.2048,55.2708],
    'deira':[25.2697,55.3094],'jumeirah':[25.2084,55.2414],'bur dubai':[25.2532,55.2988],
    'abu dhabi center':[24.4539,54.3773],'khalidiyah':[24.4720,54.3500],
    // Australia
    'sydney cbd':[-33.8688,151.2093],'surry hills':[-33.8882,151.2101],
    'melbourne cbd':[-37.8136,144.9631],'fitzroy':[-37.7996,144.9779],
    'brisbane cbd':[-27.4698,153.0251],
    // Canada
    'downtown toronto':[43.6532,-79.3832],'the annex':[43.6677,-79.4073],
    'downtown vancouver':[49.2827,-123.1207],'kitsilano':[49.2677,-123.1622],
    'downtown calgary':[51.0447,-114.0719],
};

function hexGrid(lat, lng, radiusKm=1.2, rings=2) {
    const pts     = [{ lat, lng }];
    const latDeg  = radiusKm / 111.0;
    const lngDeg  = radiusKm / (111.0 * Math.cos(lat * Math.PI / 180));
    for (let ring=1; ring<=rings; ring++) {
        for (let i=0; i<6; i++) {
            for (let j=0; j<ring; j++) {
                const ang = (i*60+30)*Math.PI/180;
                const off = ring - j;
                pts.push({
                    lat: lat + latDeg*ring*Math.sin(ang) + latDeg*off*Math.sin(ang+Math.PI/3),
                    lng: lng + lngDeg*ring*Math.cos(ang) + lngDeg*off*Math.cos(ang+Math.PI/3),
                });
            }
        }
    }
    return pts;
}

function getAreaCenter(area) {
    return AREA_COORDS[area.toLowerCase().trim()] || null;
}

// ════════════════════════════════════════════════════════════════
//   WORLD DATA
// ════════════════════════════════════════════════════════════════
const WORLD = {
    '🇮🇳 India': {
        'Tamil Nadu':{'Chennai':['T Nagar','Anna Nagar','Velachery','Adyar','Nungambakkam','Tambaram','Porur','Chromepet','Perambur','Mogappair','Kodambakkam','Mylapore','Royapettah','Egmore','Ambattur'],'Coimbatore':['RS Puram','Gandhipuram','Saibaba Colony','Peelamedu','Singanallur','Townhall','Vadavalli'],'Madurai':['Anna Nagar','KK Nagar','Bypass Road','Iyer Bungalow','SS Colony','Tallakulam'],'Salem':['Fairlands','Junction','Five Roads','Gugai','Shevapet'],'Tiruchirappalli':['Thillai Nagar','KK Nagar','Cantonment','Srirangam','Ariyamangalam']},
        'Maharashtra':{'Mumbai':['Andheri','Bandra','Dadar','Borivali','Malad','Powai','Thane','Kurla','Goregaon','Kandivali','Vile Parle','Santacruz','Juhu','Chembur','Mulund'],'Pune':['Koregaon Park','Kothrud','Wakad','Baner','Hadapsar','Viman Nagar','Hinjewadi','Aundh','Kharadi','Pimple Saudagar'],'Nagpur':['Dharampeth','Sitabuldi','Sadar','Ramdaspeth','Bajaj Nagar','Pratap Nagar'],'Nashik':['Gangapur Road','College Road','Cidco','Indira Nagar','Panchavati'],'Aurangabad':['Cidco','Garkheda','Osmanpura','Cantonment','TV Centre']},
        'Karnataka':{'Bangalore':['Koramangala','Indiranagar','Whitefield','HSR Layout','JP Nagar','Jayanagar','Electronic City','Marathahalli','Rajajinagar','Malleshwaram','Yelahanka','Bannerghatta Road','Sarjapur Road','Hebbal','RT Nagar'],'Mysore':['Vijayanagar','Kuvempunagar','Saraswathipuram','Jayalakshmipuram','Gokulam'],'Hubli':['Vidyanagar','Keshwapur','Deshpande Nagar','Gokul Road'],'Mangalore':['Hampankatta','Bejai','Kadri','Kankanady','Attavar'],'Belgaum':['Camp','Tilakwadi','Shahapur','Khanapur Road']},
        'Telangana':{'Hyderabad':['Banjara Hills','Jubilee Hills','Madhapur','Kondapur','Kukatpally','Ameerpet','Begumpet','Himayatnagar','Secunderabad','LB Nagar','Dilsukhnagar','Uppal','Gachibowli','Miyapur','Nizampet'],'Warangal':['Hanamkonda','Kazipet','Subedari','Lashkar Bazar'],'Karimnagar':['Mukarampura','Jagtial Road','Godavarikhani'],'Nizamabad':['Beside Bus Stand','Dichpally Road','Armoor Road'],'Khammam':['Wyra Road','Kothagudem','Yellandu Road']},
        'Andhra Pradesh':{'Visakhapatnam':['Dwaraka Nagar','MVP Colony','Gajuwaka','Madhurawada','Siripuram','Rushikonda','Seethammadhara','NAD Junction'],'Vijayawada':['Benz Circle','MG Road','Governorpet','Labbipet','Patamata','Moghalrajpuram','Auto Nagar'],'Guntur':['Brodipet','Arundelpet','Nagarampalem','Kothapet'],'Nellore':['Pogathota','Magunta Layout','Santhapet','Grand Bazaar'],'Kakinada':['Main Road','Suryaraopeta','Gandhinagar','Ramanayyapeta']},
        'Gujarat':{'Ahmedabad':['CG Road','Vastrapur','Navrangpura','Bopal','Satellite','Prahlad Nagar','Maninagar','Ghatlodia','Chandkheda','Motera'],'Surat':['Adajan','Vesu','Piplod','Athwa','City Light','Pal','Katargam','Varachha','Udhna'],'Vadodara':['Alkapuri','Gotri','Manjalpur','Waghodia Road','Fatehgunj'],'Rajkot':['Kalawad Road','Yagnik Road','Bhaktinagar','Raiya Road','University Road'],'Gandhinagar':['Sector 7','Sector 11','Sector 16','Sector 21','Sector 28']},
        'Delhi NCR':{'Delhi':['Connaught Place','Lajpat Nagar','Karol Bagh','Dwarka','Rohini','Pitampura','Janakpuri','Saket','Vasant Kunj','Preet Vihar','Mayur Vihar','Shahdara','Nehru Place'],'Noida':['Sector 18','Sector 62','Sector 137','Sector 50','Sector 76'],'Gurgaon':['DLF Phase 1','Sohna Road','MG Road','Golf Course Road','Sector 14'],'Faridabad':['NIT','Sector 16','Old Faridabad','Ballabhgarh'],'Ghaziabad':['Indirapuram','Vaishali','Kaushambi','Raj Nagar']},
        'Kerala':{'Kochi':['MG Road','Kakkanad','Edapally','Aluva','Kaloor','Vyttila','Tripunithura','Kalamassery'],'Thiruvananthapuram':['Palayam','Kowdiar','Pattom','Kesavadasapuram','Kazhakuttam'],'Kozhikode':['SM Street','Mavoor Road','Nadakkavu','Westhill'],'Thrissur':['Round South','Punkunnam','Ollur','Ayyanthole'],'Kannur':['Fort Road','Thavakkara','Chala','Puthiyatheru']},
        'Punjab':{'Ludhiana':['Model Town','BRS Nagar','Sarabha Nagar','Civil Lines','Rajguru Nagar'],'Amritsar':['Lawrence Road','Ranjit Avenue','Queens Road','GT Road','Majitha Road'],'Chandigarh':['Sector 17','Sector 22','Sector 34','Sector 43','Sector 9'],'Jalandhar':['Model Town','Lajpat Nagar','Civil Lines','Guru Nanak Mission Chowk'],'Patiala':['Model Town','Leela Bhawan','Tripuri','Urban Estate']},
        'Rajasthan':{'Jaipur':['Vaishali Nagar','Malviya Nagar','C Scheme','Mansarovar','Tonk Road','Ajmer Road','Jagatpura','Pratap Nagar'],'Jodhpur':['Sardarpura','Shastri Nagar','Pal Road','Residency Road'],'Udaipur':['Hiran Magri','Shobhagpura','Sukhadia Circle','Fatehpura'],'Kota':['Talwandi','Vigyan Nagar','Dadabari','Mahaveer Nagar'],'Ajmer':['Vaishali Nagar','Naya Bazar','Station Road']},
        'Uttar Pradesh':{'Lucknow':['Hazratganj','Gomti Nagar','Indira Nagar','Aliganj','Alambagh','Mahanagar','Rajajipuram'],'Agra':['Sanjay Place','Kamla Nagar','Shahganj','Civil Lines'],'Varanasi':['Sigra','Lanka','Cantt','Lahurabir','Mahmoorganj'],'Kanpur':['Civil Lines','Arya Nagar','Swaroop Nagar','Kidwai Nagar'],'Allahabad':['Civil Lines','George Town','Katra','Mumfordganj']},
        'West Bengal':{'Kolkata':['Park Street','Salt Lake','Ballygunge','Howrah','Dum Dum','Behala','Jadavpur','New Town','Rajarhat'],'Howrah':['Shibpur','Golabari','Santragachi','Liluah'],'Durgapur':['City Centre','Benachity','Bidhannagar','Steel Township'],'Siliguri':['Sevoke Road','Pradhan Nagar','Hill Cart Road','Matigara'],'Asansol':['GT Road','Burnpur','Raniganj','Kulti']},
        'Madhya Pradesh':{'Bhopal':['New Market','MP Nagar','Arera Colony','Kolar Road','Habibganj'],'Indore':['Vijay Nagar','Palasia','MG Road','Bhawarkuan','Scheme 54'],'Jabalpur':['Wright Town','Napier Town','Civil Lines','Adhartal'],'Gwalior':['City Centre','Lashkar','Morar','Thatipur'],'Ujjain':['Freeganj','Madhav Nagar','Mahakal Road']},
        'Haryana':{'Gurgaon':['DLF Phase 1','Sohna Road','MG Road','Golf Course Road','Sector 14'],'Faridabad':['NIT','Sector 16','Old Faridabad','Ballabhgarh'],'Hisar':['Urban Estate','Red Square Market','Model Town'],'Rohtak':['Civil Lines','Model Town','Sector 14'],'Panipat':['Model Town','GT Road','Sector 11']},
        'Odisha':{'Bhubaneswar':['Patia','Sahid Nagar','CDA','Nayapalli','Chandrasekharpur'],'Cuttack':['Badambadi','Buxi Bazaar','CDA','Chandini Chowk'],'Rourkela':['Udit Nagar','Civil Township','Sector 19','Chhend'],'Puri':['Grand Road','Marine Drive','CT Road'],'Sambalpur':['Ainthapali','Budharaja','Nayapara']},
        'Uttarakhand':{'Dehradun':['Rajpur Road','Patel Nagar','Ballupur','Vasant Vihar'],'Haridwar':['Jwalapur','Ranipur More','SIDCUL'],'Roorkee':['Civil Lines','IIT Roorkee Area','Railway Road'],'Nainital':['Mall Road','Tallital','Mallital'],'Rishikesh':['Laxman Jhula','Ram Jhula','Muni Ki Reti']},
        'Jharkhand':{'Ranchi':['Lalpur','Harmu','Kanke Road','Ashok Nagar'],'Jamshedpur':['Bistupur','Sakchi','Telco','Sonari'],'Dhanbad':['Bank More','Hirapur','Jharia','Katras'],'Bokaro':['Sector 4','City Centre','Chas'],'Hazaribagh':['Town','Annanda Chowk','Canary Hill']},
        'Chhattisgarh':{'Raipur':['Shankar Nagar','Civil Lines','Telibandha','Pandri'],'Bilaspur':['Vyapar Vihar','Torwa','Sadar Bazar'],'Durg':['Nehru Nagar','Smriti Nagar','Supela'],'Bhilai':['Sector 1','Sector 6','Civic Centre'],'Korba':['Transport Nagar','NTPC Colony','City Kotwali']},
        'Assam':{'Guwahati':['Dispur','Paltan Bazaar','Bhangagarh','Zoo Road','GS Road'],'Silchar':['Premtala','Rangirkhari','Link Road'],'Dibrugarh':['Chowkidinghee','AT Road','Graham Bazaar'],'Jorhat':['AT Road','MG Road','Gar Ali'],'Nagaon':['Haibargaon','RG Baruah Road']},
        'Himachal Pradesh':{'Shimla':['Mall Road','Sanjauli','Vikas Nagar','Khalini'],'Manali':['Mall Road','Old Manali','Aleo'],'Dharamshala':['McLeod Ganj','Kotwali Bazaar','Sidhpur'],'Solan':['The Mall','Salogra','Chambaghat'],'Mandi':['Indira Market','Purani Mandi','Nehru Park']},
        'Goa':{'Panaji':['MG Road','Fontainhas','Miramar','Dona Paula'],'Margao':['Monte Hill','Fatorda','Aquem'],'Vasco':['Baina','Chicalim','Vaddem'],'Mapusa':['Friday Market','Angod','Durbhat'],'Ponda':['Tisk','Curti','Farmagudi']},
        'Jammu & Kashmir':{'Jammu':['Gandhi Nagar','Bakshi Nagar','Trikuta Nagar','Raghunath Bazaar'],'Srinagar':['Lal Chowk','Rajbagh','Jawahar Nagar','Hyderpora'],'Anantnag':['Town','Shahabad','Achabal'],'Baramulla':['Town','Sopore','Pattan'],'Udhampur':['Town','Chenab Nagar']},
        'Puducherry':{'Puducherry':['MG Road','Anna Salai','Lawspet','Villianur'],'Karaikal':['Town','Neravy','Tirumalairayanpattinam'],'Mahe':['Town','Palloor'],'Yanam':['Town'],'Ozhukarai':['Town']},
    },
    '🇺🇸 USA': {
        'New York':{'New York City':['Manhattan','Brooklyn','Queens','Bronx','Midtown Manhattan','Upper East Side','Williamsburg','Astoria','Flushing'],'Buffalo':['Downtown','Elmwood Village','Allentown'],'Rochester':['Downtown','Park Avenue','Brighton']},
        'California':{'Los Angeles':['Downtown Los Angeles','Hollywood','Santa Monica','Beverly Hills','West Hollywood','Silver Lake','Koreatown','Culver City','Pasadena'],'San Francisco':['Downtown','Mission District','SoMa','Castro','Noe Valley','Richmond','North Beach'],'San Diego':['Downtown','Gaslamp Quarter','North Park','Hillcrest','La Jolla']},
        'Texas':{'Houston':['Downtown Houston','Midtown','Montrose','The Heights','Galleria'],'Dallas':['Downtown Dallas','Uptown','Oak Lawn','Deep Ellum','Plano'],'Austin':['Downtown Austin','South Congress','East Austin','Hyde Park']},
        'Illinois':{'Chicago':['Downtown Chicago','Lincoln Park','Wicker Park','Logan Square','Hyde Park','Evanston']},
        'Florida':{'Miami':['Downtown Miami','South Beach','Brickell','Wynwood','Coral Gables'],'Orlando':['Downtown Orlando','Winter Park','Kissimmee'],'Tampa':['Downtown Tampa','Hyde Park','Ybor City']},
    },
    '🇬🇧 UK': {
        'England':{'London':['Central London','Shoreditch','Canary Wharf','City of London','Camden','Islington','Hackney','Brixton','Notting Hill','Kensington','Chelsea'],'Manchester':['Central Manchester','Northern Quarter','Didsbury','Chorlton','Salford'],'Birmingham':['Central Birmingham','Digbeth','Moseley','Edgbaston'],'Leeds':['City Centre','Headingley','Chapel Allerton'],'Bristol':['City Centre','Clifton','Stokes Croft']},
        'Scotland':{'Edinburgh':['Old Town','New Town','Leith','Morningside'],'Glasgow':['City Centre','West End','Merchant City']},
        'Wales':{'Cardiff':['City Centre','Pontcanna','Roath','Canton']},
    },
    '🇦🇪 UAE': {
        'Dubai':{'Dubai':['Dubai Marina','Downtown Dubai','Deira','Jumeirah','Bur Dubai','JLT','Business Bay','DIFC','Karama','Al Barsha','Mirdif']},
        'Abu Dhabi':{'Abu Dhabi':['Abu Dhabi Center','Khalidiyah','Al Reem Island','Yas Island','Saadiyat Island','Mussafah']},
        'Sharjah':{'Sharjah':['Al Nahda','Al Khan','Muwaileh','Al Qasimia']},
    },
    '🇦🇺 Australia': {
        'New South Wales':{'Sydney':['Sydney CBD','Surry Hills','Newtown Sydney','Glebe','Paddington','Bondi','Manly','Parramatta','Chatswood']},
        'Victoria':{'Melbourne':['Melbourne CBD','Fitzroy','Collingwood','Richmond','St Kilda','South Yarra','Carlton']},
        'Queensland':{'Brisbane':['Brisbane CBD','South Brisbane','Fortitude Valley','New Farm','West End'],'Gold Coast':['Surfers Paradise','Broadbeach','Burleigh Heads']},
        'Western Australia':{'Perth':['Perth CBD','Fremantle','Northbridge','Subiaco']},
    },
    '🇨🇦 Canada': {
        'Ontario':{'Toronto':['Downtown Toronto','The Annex','Kensington Market','Little Italy','Leslieville','North York','Mississauga'],'Ottawa':['Downtown Ottawa','Byward Market','Glebe','Westboro']},
        'British Columbia':{'Vancouver':['Downtown Vancouver','Kitsilano','Commercial Drive','Mount Pleasant','Richmond','Burnaby']},
        'Alberta':{'Calgary':['Downtown Calgary','Kensington','Mission','Inglewood','Beltline'],'Edmonton':['Downtown Edmonton','Whyte Avenue','Old Strathcona']},
        'Quebec':{'Montreal':['Downtown Montreal','Plateau-Mont-Royal','Mile End','Old Montreal','Rosemont']},
    },
};

const ALL_CATEGORIES = [
    'dentist','dental clinic','skin clinic','cosmetic clinic',
    'physiotherapy center','veterinary clinic','gym','yoga studio',
    'zumba classes','pilates studio','weight loss center','salon',
    'beauty parlour','spa','eye clinic','pharmacy','diagnostic center',
    'AC repair','electrician','plumber','mobile repair','car repair',
];

// ════════════════════════════════════════════════════════════════
//   FILE HELPERS
// ════════════════════════════════════════════════════════════════
const RESULTS_DIR   = 'results';
const RESUME_FILE   = 'resume_progress.json';
const HISTORY_FILE  = 'scrape_history.log';
const COMBINED_FILE = join(RESULTS_DIR, 'ALL_RESULTS_COMBINED');

function todayDate() {
    const d=new Date(), months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${String(d.getDate()).padStart(2,'0')} ${months[d.getMonth()]} ${d.getFullYear()}`;
}
function formatTime(ms) {
    const s=Math.floor(ms/1000),m=Math.floor(s/60),h=Math.floor(m/60);
    if(h>0) return `${h}h ${m%60}m`; if(m>0) return `${m}m ${s%60}s`; return `${s}s`;
}
function dedupeKey(title,city) { return `${(title||'').toLowerCase().trim()}|${(city||'').toLowerCase().trim()}`; }
async function getText(page,sel) { try { const el=await page.$(sel); return (await el?.textContent())?.trim()||null; } catch { return null; } }
function isBigChain(name) { const n=(name||'').toLowerCase(); return BIG_CHAINS.some(c=>n.includes(c)); }
function loadResume() { try { if(existsSync(RESUME_FILE)) return JSON.parse(readFileSync(RESUME_FILE,'utf8')); } catch {} return null; }
function saveResume(data) { try { writeFileSync(RESUME_FILE,JSON.stringify(data,null,2)); } catch {} }
function appendHistory(msg) {
    try {
        const old=existsSync(HISTORY_FILE)?readFileSync(HISTORY_FILE,'utf8'):'';
        writeFileSync(HISTORY_FILE,old+`[${new Date().toLocaleString()}] ${msg}\n`);
    } catch {}
}

async function saveRealExcel(profiles, filePath) {
    try {
        if (!profiles?.length) return;
        const wb = new ExcelJS.Workbook();
        wb.creator = 'GBP Fix Scraper v2';
        wb.created = new Date();
        const ws = wb.addWorksheet('Results', { views:[{ state:'frozen', ySplit:1 }] });

        ws.columns = [
            { header:'City',           key:'city',         width:16 },
            { header:'State / Region', key:'state',        width:18 },
            { header:'Country',        key:'country',      width:14 },
            { header:'Category',       key:'categoryName', width:22 },
            { header:'Business Name',  key:'title',        width:40 },
            { header:'Phone',          key:'phone',        width:16 },
            { header:'Rating',         key:'totalScore',   width:10 },
            { header:'Reviews',        key:'reviewCount',  width:10 },
            { header:'Website',        key:'website',      width:36 },
            { header:'Latitude',       key:'latitude',     width:14 },
            { header:'Longitude',      key:'longitude',    width:14 },
            { header:'Google Maps URL',key:'googleMapsUrl',width:20 },
            { header:'Scraped Date',   key:'scrapedDate',  width:14 },
        ];

        // Bold blue header
        ws.getRow(1).eachCell(cell => {
            cell.fill      = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF1A3C5E' } };
            cell.font      = { bold:true, color:{ argb:'FFFFFFFF' }, size:11 };
            cell.alignment = { vertical:'middle', horizontal:'center' };
            cell.border    = { bottom:{ style:'medium', color:{ argb:'FF2196F3' } } };
        });
        ws.getRow(1).height = 22;

        profiles.forEach((p, idx) => {
            const row = ws.addRow({
                city:p.city||'', state:p.state||'', country:p.country||'',
                categoryName:p.categoryName||'', title:p.title||'', phone:p.phone||'',
                totalScore:p.totalScore||'', reviewCount:p.reviewCount||'',
                website:p.website||'', latitude:p.latitude||'', longitude:p.longitude||'',
                googleMapsUrl:p.googleMapsUrl||'', scrapedDate:p.scrapedDate||'',
            });
            const bg = idx%2===0 ? 'FFF0F4F8' : 'FFFFFFFF';
            row.eachCell(cell => {
                cell.fill      = { type:'pattern', pattern:'solid', fgColor:{ argb:bg } };
                cell.alignment = { vertical:'middle' };
            });
            if (p.googleMapsUrl) {
                const c = row.getCell('googleMapsUrl');
                c.value = { text:'Open in Maps', hyperlink:p.googleMapsUrl };
                c.font  = { color:{ argb:'FF1565C0' }, underline:true };
            }
            if (p.website) {
                try {
                    const c = row.getCell('website');
                    c.value = { text:p.website.replace(/^https?:\/\//,''), hyperlink:p.website };
                    c.font  = { color:{ argb:'FF1565C0' }, underline:true };
                } catch {}
            }
            ['totalScore','reviewCount','latitude','longitude'].forEach(k => {
                row.getCell(k).alignment = { horizontal:'center', vertical:'middle' };
            });
            row.height = 18;
        });

        ws.autoFilter = { from:'A1', to:'M1' };
        await wb.xlsx.writeFile(filePath+'.xlsx');
        log.info(`Saved: ${filePath}.xlsx (${profiles.length} profiles)`);
    } catch (e) { log.warning(`Excel error: ${e.message}`); }
}

function saveCSV(profiles, filePath) {
    try {
        if (!profiles?.length) return;
        const h=Object.keys(profiles[0]);
        writeFileSync(filePath+'.csv',[h.join(','),...profiles.map(r=>h.map(k=>`"${String(r[k]||'').replace(/"/g,'""')}"`).join(','))].join('\n'),'utf8');
    } catch(e) { log.warning(`CSV error: ${e.message}`); }
}

async function saveFiles(profiles, city, category) {
    try {
        if (!existsSync(RESULTS_DIR)) mkdirSync(RESULTS_DIR);
        const p = join(RESULTS_DIR,`${city.replace(/[^a-z0-9]/gi,'_')}_${category.replace(/[^a-z0-9]/gi,'_')}_${todayDate().replace(/ /g,'-')}`);
        saveCSV(profiles, p);
        await saveRealExcel(profiles, p);
    } catch(e) { log.warning(`Save error: ${e.message}`); }
}

async function saveCombined(all) {
    try {
        if (!existsSync(RESULTS_DIR)) mkdirSync(RESULTS_DIR);
        saveCSV(all, COMBINED_FILE);
        await saveRealExcel(all, COMBINED_FILE);
    } catch(e) { log.warning(`Combined error: ${e.message}`); }
}

// ════════════════════════════════════════════════════════════════
//   MENU HELPERS
// ════════════════════════════════════════════════════════════════
function prompt(q) {
    const rl = readline.createInterface({ input:process.stdin, output:process.stdout });
    return new Promise(r => rl.question(q, a => { rl.close(); r(a.trim()); }));
}

async function menu(title, options, multi=false) {
    console.log(`\n${'═'.repeat(52)}\n  ${title}\n${'═'.repeat(52)}`);
    options.forEach((o,i) => console.log(`  ${String(i+1).padStart(2)}. ${o}`));
    if (multi) {
        console.log(`  ${String(options.length+1).padStart(2)}. ALL`);
        console.log('\n  Type numbers with comma (e.g. 1,3,5) or ALL');
        const ans = await prompt('\n  Your choice: ');
        if (ans.toUpperCase()==='ALL' || ans===String(options.length+1)) return [...options];
        return ans.split(',').map(n=>parseInt(n.trim())-1).filter(n=>n>=0&&n<options.length).map(i=>options[i]);
    } else {
        const ans = await prompt('\n  Your choice (number): ');
        return options[parseInt(ans)-1] || options[0];
    }
}

// ════════════════════════════════════════════════════════════════
//   MAIN
// ════════════════════════════════════════════════════════════════
await Actor.init();
if (!existsSync(RESULTS_DIR)) mkdirSync(RESULTS_DIR);

console.log('\n' + '═'.repeat(52));
console.log('   GOOGLE MAPS SCRAPER v2.0 — GBP Fix Tool');
console.log('   Multi-Country | Geo-Grid | Real Excel');
console.log('═'.repeat(52));
console.log('\n  1. NEW SCRAPE    — Start fresh');
console.log('  2. CONTINUE      — Resume from where stopped');
console.log('  3. RE-SCRAPE     — Redo a specific area');

const startOpt = await prompt('\n  Type 1, 2 or 3: ');

let completedWork = new Set();
let seenProfiles  = new Set();
let allSaved      = [];

if (startOpt === '2' || startOpt === '3') {
    const r = loadResume();
    if (r) {
        completedWork = new Set(r.completedWork||[]);
        seenProfiles  = new Set(r.seenProfiles||[]);
        allSaved      = r.allSaved||[];
        console.log(`\n  ✅ Resume loaded! ${allSaved.length} profiles previously saved.`);
        if (r.inProgress) console.log(`  In progress: ${r.inProgress}`);
    } else {
        console.log('\n  No resume found. Starting fresh...');
    }
}

// ─── COUNTRY SELECTION ───────────────────────────────────────
const countryNames = Object.keys(WORLD);
const selCountries = await menu('STEP 1: Select COUNTRY', countryNames, true);

// Custom country
console.log('\n  ✏️  Add a CUSTOM country not in list?');
const customCountry = await prompt('  Country name (or ENTER to skip): ');
if (customCountry.trim()) {
    const cCity  = await prompt('  City name: ');
    const cAreas = await prompt('  Areas (comma separated): ');
    WORLD[customCountry.trim()] = { 'Custom':{ [cCity.trim()]: cAreas.split(',').map(a=>a.trim()).filter(Boolean) } };
    selCountries.push(customCountry.trim());
}

// ─── REGION/CITY/AREA SELECTION ──────────────────────────────
const selAreas = [];
for (const country of selCountries) {
    const cData    = WORLD[country];
    const rNames   = Object.keys(cData);
    const selRegs  = rNames.length===1 ? rNames : await menu(`Regions in ${country}`, rNames, true);

    for (const region of selRegs) {
        const cityNames = Object.keys(cData[region]);
        const selCities = await menu(`Cities in ${region}`, cityNames, true);

        // Custom city
        console.log(`\n  ✏️  Add a CUSTOM city in ${region}?`);
        const customCity = await prompt('  City name (or ENTER to skip): ');
        if (customCity.trim()) {
            const ca = await prompt('  Areas (comma separated): ');
            cData[region][customCity.trim()] = ca.split(',').map(a=>a.trim()).filter(Boolean);
            selCities.push(customCity.trim());
        }

        for (const city of selCities) {
            const areaList = cData[region][city]||[];
            const extraOpt = [...areaList, '✏️  TYPE CUSTOM AREA'];
            const chosen   = await menu(`Areas in ${city}`, extraOpt, true);
            let finalAreas = chosen.filter(a=>!a.includes('TYPE CUSTOM AREA'));
            if (chosen.some(a=>a.includes('TYPE CUSTOM AREA'))) {
                const ca = await prompt('  Custom areas (comma separated): ');
                ca.split(',').forEach(a=>{ if(a.trim()) finalAreas.push(a.trim()); });
            }
            selAreas.push({ country:country.replace(/^\S+\s/,''), region, city, areas:finalAreas });
        }
    }
}

// ─── CATEGORIES ──────────────────────────────────────────────
const selCategories = await menu('STEP 2: Select CATEGORIES', ALL_CATEGORIES, true);
console.log('\n  ✏️  Add a CUSTOM category?');
const customCat = await prompt('  Category (or ENTER to skip): ');
if (customCat.trim()) selCategories.push(customCat.trim());

// ─── FILTERS ─────────────────────────────────────────────────
console.log('\n' + '═'.repeat(52) + '\n  FILTERS\n' + '═'.repeat(52));
const phoneFilter   = await menu('Phone filter',   ['without phone only','with phone only','any (both)']);
const websiteFilter = await menu('Website filter', ['any (both)','without website','with website']);
const filters = {
    phone  : phoneFilter.includes('without') ? 'without' : phoneFilter.includes('with phone') ? 'with' : 'any',
    website: websiteFilter.includes('without') ? 'without' : websiteFilter.includes('with website') ? 'with' : 'any',
};

// ─── GEO-GRID ────────────────────────────────────────────────
console.log('\n' + '═'.repeat(52) + '\n  GEO-GRID SEARCH\n' + '═'.repeat(52));
console.log('  Geo-grid places GPS search points over each area');
console.log('  giving 5-10x more results than text search only.\n');
console.log('  1. Enable  (recommended — thorough coverage)');
console.log('  2. Disable (faster but fewer results)');
const useGeoGrid = (await prompt('\n  Your choice (number): ')) !== '2';

// ─── TARGET & TIME ────────────────────────────────────────────
console.log('\n' + '═'.repeat(52) + '\n  TARGET & TIME\n' + '═'.repeat(52));
const targetListings = parseInt(await prompt('\n  How many profiles? (e.g. 50): '))||50;
const timeOpt        = await menu('Time limit', ['10 minutes','30 minutes','60 minutes','2 hours','No limit']);
const maxMinutes     = timeOpt.includes('No')?999999:timeOpt.includes('10')?10:timeOpt.includes('30')?30:timeOpt.includes('60')?60:120;

console.log('\n' + '═'.repeat(52));
console.log('  SUMMARY');
console.log('═'.repeat(52));
console.log(`  Countries  : ${[...new Set(selAreas.map(a=>a.country))].join(', ')}`);
console.log(`  Cities     : ${[...new Set(selAreas.map(a=>a.city))].join(', ')}`);
console.log(`  Categories : ${selCategories.join(', ')}`);
console.log(`  Target     : ${targetListings} | Time: ${maxMinutes===999999?'No limit':maxMinutes+'min'}`);
console.log(`  Phone      : ${filters.phone} | Website: ${filters.website}`);
console.log(`  Geo-grid   : ${useGeoGrid?'✅ ON':'❌ OFF'}`);
console.log('═'.repeat(52));

if ((await prompt('\n  Start? (Y/N): ')).toUpperCase() !== 'Y') { console.log('Cancelled.'); process.exit(0); }

// ─── RUN STATE ───────────────────────────────────────────────
const dataset    = await Actor.openDataset();
const START_TIME = Date.now();
const MAX_MS     = maxMinutes * 60 * 1000;
const isTimeUp   = () => Date.now()-START_TIME >= MAX_MS;
const elapsed    = () => formatTime(Date.now()-START_TIME);

let totalSaved    = allSaved.length;
let totalVisited  = 0;
let skippedCount  = 0;
let skippedDupe   = 0;
const cityProfiles  = {};
const catSummary    = {};
const placeUrls     = new Set();

if (allSaved.length>0) await saveCombined(allSaved);

// ─── BLOCK DETECTION ────────────────────────────────────────
let blockCount = 0;
async function checkBlocked(page) {
    try {
        const body = await page.evaluate(()=>document.title+' '+document.body.innerText.substring(0,400)).catch(()=>'');
        if (/unusual traffic|captcha|robot|sorry/i.test(body) || page.url().includes('sorry/index')) {
            blockCount++;
            log.warning(`⚠️ Block #${blockCount} detected. Pausing 3 minutes...`);
            console.log('\n  ⚠️  Google rate limit hit. Auto-pausing 3 minutes...');
            await sleep(180000);
            return true;
        }
    } catch {}
    return false;
}

// ─── BROWSER BASE ARGS ───────────────────────────────────────
const BASE_ARGS = [
    '--no-sandbox','--disable-setuid-sandbox','--disable-dev-shm-usage',
    '--lang=en-US,en','--disable-blink-features=AutomationControlled',
];

// ─── CRAWLER FACTORIES ───────────────────────────────────────
function makeSearchCrawler() {
    const ua=nextUA(), vp=nextViewport();
    return new PlaywrightCrawler({
        headless:true, useSessionPool:false, persistCookiesPerSession:false,
        maxConcurrency:2, requestHandlerTimeoutSecs:60, navigationTimeoutSecs:30,
        launchContext:{ launchOptions:{ args:[...BASE_ARGS,`--window-size=${vp.width},${vp.height}`] } },
        browserPoolOptions:{ maxOpenPagesPerBrowser:1 },
        preNavigationHooks:[async ({page})=>{
            await page.setViewportSize(vp);
            await page.setExtraHTTPHeaders({'Accept-Language':'en-US,en;q=0.9'});
            await page.evaluate(ua=>{
                Object.defineProperty(navigator,'webdriver',{get:()=>undefined});
                Object.defineProperty(navigator,'userAgent',{get:()=>ua});
            },ua);
        }],
        async requestHandler({page,request}) {
            if (isTimeUp()||totalSaved>=targetListings) return;
            await sleep(randDelay(600,1200));
            if (await checkBlocked(page)) return;
            await page.locator('button:has-text("Accept all"),button:has-text("Reject all")').first().click({timeout:3000}).catch(()=>{});
            await sleep(400);
            const hasFeed = await page.$(SEL.feed).catch(()=>null);
            if (!hasFeed) {
                const ct = await page.evaluate(()=>document.querySelectorAll('a[href*="/maps/place/"]').length);
                if (ct===0) return;
            }
            const found=new Set(); let noNew=0;
            for (let i=0;i<15;i++) {
                const links = await page.evaluate(()=>Array.from(document.querySelectorAll('a[href*="/maps/place/"]')).map(a=>a.href.split('?')[0]+'?hl=en'));
                let newCount=0;
                for (const link of links) { if (!found.has(link)&&!seenProfiles.has(link)) { found.add(link); newCount++; } }
                if (newCount===0) { noNew++; if(noNew>=3) break; } else noNew=0;
                await page.evaluate(()=>{ const f=document.querySelector('div[role="feed"]')||document.querySelector('div.m6QErb'); if(f){f.scrollTop+=3000;f.scrollBy(0,3000);} });
                await sleep(randDelay(800,1400));
            }
            const {area,city}=request.userData;
            let added=0;
            for (const url of found) { if (!placeUrls.has(url)) { placeUrls.add(url); added++; } }
            if (added>0) log.info(`${city}/${area} → +${added} URLs (total:${placeUrls.size})`);
        },
        failedRequestHandler({request}){ log.warning(`Search failed: ${request.url}`); },
    });
}

function makeDetailCrawler() {
    const ua=nextUA(), vp=nextViewport();
    return new PlaywrightCrawler({
        headless:true, useSessionPool:false, persistCookiesPerSession:false,
        maxConcurrency:2, requestHandlerTimeoutSecs:45, navigationTimeoutSecs:20,
        launchContext:{ launchOptions:{ args:[...BASE_ARGS,`--window-size=${vp.width},${vp.height}`] } },
        browserPoolOptions:{ maxOpenPagesPerBrowser:1 },
        preNavigationHooks:[async ({page})=>{
            await page.setViewportSize(vp);
            await page.setExtraHTTPHeaders({'Accept-Language':'en-US,en;q=0.9'});
            await page.evaluate(ua=>{
                Object.defineProperty(navigator,'webdriver',{get:()=>undefined});
                Object.defineProperty(navigator,'userAgent',{get:()=>ua});
            },ua);
        }],
        async requestHandler({page,request}) {
            if (isTimeUp()||totalSaved>=targetListings) return;
            await sleep(randDelay(400,800));
            await page.locator('button:has-text("Accept all"),button:has-text("Reject all")').first().click({timeout:3000}).catch(()=>{});
            if (await checkBlocked(page)) return;
            try { await page.waitForSelector(SEL.title,{timeout:12000}); } catch { return; }
            await sleep(randDelay(500,900));
            await page.waitForTimeout(500).catch(()=>{});
            totalVisited++;

            const {city,state,country,category,workKey}=request.userData;

            // Phone
            const phoneEl = await page.$(SEL.phone).catch(()=>null);
            let phoneNumber='';
            if (phoneEl) {
                phoneNumber=(await phoneEl.textContent().catch(()=>''))?.trim()||'';
                const hasPhone=phoneNumber&&!/add phone|add a phone/i.test(phoneNumber);
                if (!hasPhone) phoneNumber='';
                if (filters.phone==='without'&&hasPhone) { skippedCount++; return; }
                if (filters.phone==='with'&&!hasPhone)   { skippedCount++; return; }
            } else {
                if (filters.phone==='with') { skippedCount++; return; }
            }

            // Closed check
            const bodyText=await page.evaluate(()=>document.body.innerText).catch(()=>'');
            if (/permanently closed/i.test(bodyText)||/temporarily closed/i.test(bodyText)) { skippedCount++; return; }

            // Title
            const title=await getText(page,SEL.title);
            if (!title||isBigChain(title)) return;

            // Website
            let website='';
            try { const w=await page.$(SEL.website); if(w) website=(await w.getAttribute('href')||'').split('?')[0]; } catch {}
            if (filters.website==='without'&&website) { skippedCount++; return; }
            if (filters.website==='with'&&!website)   { skippedCount++; return; }

            // Rating + Reviews — 5 fallback methods
            let totalScore='', reviewCount='';
            try {
                const ex=await page.evaluate(()=>{
                    let rating='',reviews='';
                    // M1: aria-label star
                    const se=document.querySelector('button[aria-label*="star"],div[aria-label*="star"],span[aria-label*="star"],button[aria-label*="Star"],div[aria-label*="Star"]');
                    if(se){const l=se.getAttribute('aria-label')||'';const rm=l.match(/([\d.]+)\s*star/i);if(rm)rating=rm[1];const rv=l.replace(/,/g,'').match(/(\d+)\s*review/i);if(rv)reviews=rv[1];}
                    // M2: F7nice
                    if(!rating){const f=document.querySelector('div.F7nice span[aria-hidden="true"]');if(f)rating=f.textContent.trim();}
                    if(!reviews){const f=document.querySelector('div.F7nice span[aria-label]');if(f){const m=(f.getAttribute('aria-label')||'').replace(/,/g,'').match(/\d+/);if(m)reviews=m[0];}}
                    // M3: scan all aria-labels
                    if(!rating||!reviews){document.querySelectorAll('[aria-label]').forEach(el=>{const l=(el.getAttribute('aria-label')||'').replace(/,/g,'');if(!rating){const m=l.match(/^([\d.]+)\s*star/i)||l.match(/([\d.]+)\s*out of 5/i);if(m)rating=m[1];}if(!reviews){const m=l.match(/(\d+)\s*review/i);if(m)reviews=m[1];}});}
                    // M4: span pattern x.x
                    if(!rating){for(const s of document.querySelectorAll('span[aria-hidden="true"],div.fontDisplayLarge,span.ceNzKf')){const t=s.textContent.trim().replace(',','.');if(/^[1-5]\.\d$/.test(t)){rating=t;break;}}}
                    // M5: (123) pattern
                    if(!reviews){for(const el of document.querySelectorAll('button,span')){const t=(el.textContent||'').trim();if(/^\([\d,]+\)$/.test(t)){reviews=t.replace(/[^0-9]/g,'');break;}}}
                    return {rating,reviews};
                });
                if(ex.rating)  totalScore  = parseFloat(ex.rating.replace(',','.'));
                if(ex.reviews) reviewCount = parseInt(ex.reviews,10);
            } catch {}

            // Coordinates from URL
            let latitude='', longitude='';
            try {
                const m=page.url().match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
                if(m){ latitude=parseFloat(m[1]); longitude=parseFloat(m[2]); }
            } catch {}

            // Category
            const categoryName=await getText(page,SEL.category)||category;

            // Dedupe
            const key=dedupeKey(title,city);
            if(seenProfiles.has(key)){skippedDupe++;return;}
            seenProfiles.add(key);
            seenProfiles.add(request.url);

            const profile={
                city, state:state||'', country:country||'', categoryName, title,
                phone:phoneNumber, totalScore, reviewCount, website,
                latitude, longitude, googleMapsUrl:page.url(), scrapedDate:todayDate(),
            };

            await dataset.pushData(profile);
            allSaved.push(profile);
            totalSaved++;
            if(!cityProfiles[workKey]) cityProfiles[workKey]=[];
            cityProfiles[workKey].push(profile);
            if(!catSummary[category]) catSummary[category]=0;
            catSummary[category]++;

            // ── INSTANT SAVE every profile ──
            await saveCombined(allSaved);
            saveResume({completedWork:[...completedWork],seenProfiles:[...seenProfiles],allSaved,inProgress:workKey});
            if(cityProfiles[workKey].length%5===0) await saveFiles(cityProfiles[workKey],city,category);

            // Progress
            const pct=Math.min(100,Math.round((totalSaved/targetListings)*100));
            const rate=totalSaved/Math.max(1,(Date.now()-START_TIME)/60000);
            const eta=rate>0?formatTime(((targetListings-totalSaved)/rate)*60000):'?';
            console.log('\n'+'─'.repeat(52));
            console.log(`  Country  : ${country}`);
            console.log(`  City     : ${city}  |  Category: ${category}`);
            console.log(`  Progress : ${totalSaved}/${targetListings} (${pct}%)  |  ETA: ${eta}`);
            console.log(`  Time     : ${elapsed()}  |  Skipped: ${skippedCount}  |  Dupes: ${skippedDupe}`);
            if(totalScore) console.log(`  Rating   : ⭐ ${totalScore}  (${reviewCount} reviews)`);
            if(latitude)   console.log(`  Coords   : ${latitude}, ${longitude}`);
            if(website)    console.log(`  Website  : ${website}`);
            console.log(`  ✅  "${title}"`);
            console.log('─'.repeat(52));
        },
        failedRequestHandler({request}){ log.warning(`Detail failed: ${request.url}`); },
    });
}

// ─── MAIN SCRAPE LOOP ────────────────────────────────────────
appendHistory(`Started: ${[...new Set(selAreas.map(a=>a.city))].join(', ')} | ${selCategories.join(', ')}`);

for (const {country,region,city,areas} of selAreas) {
    if (totalSaved>=targetListings||isTimeUp()) break;
    const state=region;

    for (const category of selCategories) {
        if (totalSaved>=targetListings||isTimeUp()) break;
        const workKey=`${country}_${city}_${category}`;

        if (startOpt==='3') {
            const ans=await prompt(`\n  Re-scrape ${city} ${category}? (Y/N): `);
            if (ans.toUpperCase()!=='Y') { completedWork.add(workKey); continue; }
            completedWork.delete(workKey);
        }
        if (completedWork.has(workKey)&&startOpt!=='1') { log.info(`Skipping ${city} ${category} — done`); continue; }

        log.info(`\n▶️  ${country} / ${city} → ${category}`);

        // Phase 1 — collect URLs
        placeUrls.clear();
        let searchReqs=[];
        const seenUrls=new Set();

        for (const area of areas) {
            if (useGeoGrid) {
                const coords=getAreaCenter(area);
                if (coords) {
                    const pts=hexGrid(coords[0],coords[1],1.2,2);
                    for (const pt of pts) {
                        const u=`https://www.google.com/maps/search/${encodeURIComponent(category)}/@${pt.lat.toFixed(6)},${pt.lng.toFixed(6)},15z?hl=en`;
                        if(!seenUrls.has(u)){seenUrls.add(u);searchReqs.push({url:u,userData:{city,state,country,category,workKey,area}});}
                    }
                    log.info(`  Geo-grid: ${hexGrid(coords[0],coords[1],1.2,2).length} points for ${area}`);
                } else {
                    // fallback text search + geo center of city if known
                    const u=`https://www.google.com/maps/search/${encodeURIComponent(category+' in '+area+' '+city)}?hl=en`;
                    if(!seenUrls.has(u)){seenUrls.add(u);searchReqs.push({url:u,userData:{city,state,country,category,workKey,area}});}
                }
            } else {
                const u=`https://www.google.com/maps/search/${encodeURIComponent(category+' in '+area+' '+city)}?hl=en`;
                if(!seenUrls.has(u)){seenUrls.add(u);searchReqs.push({url:u,userData:{city,state,country,category,workKey,area}});}
            }
        }

        log.info(`  Phase 1: ${searchReqs.length} search requests`);
        const sc=makeSearchCrawler();
        await sc.run(searchReqs);
        log.info(`  Phase 1 done: ${placeUrls.size} unique URLs`);

        if (placeUrls.size===0) { completedWork.add(workKey); continue; }

        // Phase 2 — extract profiles
        const dc=makeDetailCrawler();
        await dc.run([...placeUrls].map(url=>({url,userData:{city,state,country,category,workKey}})));

        if (cityProfiles[workKey]?.length>0) {
            await saveFiles(cityProfiles[workKey],city,category);
            appendHistory(`Done: ${city} ${category} → ${cityProfiles[workKey].length} profiles`);
        }

        completedWork.add(workKey);
        await saveCombined(allSaved);
        saveResume({completedWork:[...completedWork],seenProfiles:[...seenProfiles],allSaved,inProgress:null});

        // 1-hour checkpoint
        if (Date.now()-START_TIME>=3600000) {
            console.log('\n'+'═'.repeat(52));
            console.log(`  ⏰ 1 HOUR — ${totalSaved} profiles saved.`);
            console.log('═'.repeat(52));
            const cont=await prompt('\n  Continue? (Y/N): ');
            if (cont.toUpperCase()!=='Y') { saveResume({completedWork:[...completedWork],seenProfiles:[...seenProfiles],allSaved,inProgress:null}); process.exit(0); }
        }
    }
}

// ─── FINAL ───────────────────────────────────────────────────
await saveCombined(allSaved);
saveResume({completedWork:[...completedWork],seenProfiles:[...seenProfiles],allSaved,inProgress:null});

console.log('\n'+'═'.repeat(52));
console.log('  🎉 SCRAPE FINISHED!');
console.log('═'.repeat(52));
console.log(`  Total saved  : ${totalSaved} profiles`);
console.log(`  Visited      : ${totalVisited}`);
console.log(`  Skipped      : ${skippedCount}`);
console.log(`  Duplicates   : ${skippedDupe}`);
console.log(`  Time         : ${elapsed()}`);
console.log('─'.repeat(52));
console.log('  BREAKDOWN BY CATEGORY:');
for (const [cat,count] of Object.entries(catSummary)) {
    console.log(`    ${cat.padEnd(24)}: ${count}`);
}
console.log('─'.repeat(52));
console.log('  Files saved to: results\\');
console.log('  To open:  explorer results');
console.log('═'.repeat(52));

appendHistory(`Finished: ${totalSaved} profiles | ${elapsed()}`);
await Actor.exit();
