from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()

companies = [
    (1,"Blink22","blink22.com","Web Dev","hello@blink22.com","contact","","","","","Medium","Send cold email; ask to forward to hiring manager"),
    (2,"Shakuro","shakuro.com","Web/UX","hi@shakuro.com","contact","","","","","Medium","Send cold email"),
    (3,"ExpandX","expandx.ae","Web Dev","info@expandx.ae","contact","","","","","High","Dubai-based"),
    (4,"Digital Gravity","digitalgravity.ae","Web Dev","discover@digitalgravity.ae","contact","","","","","High","Dubai-based; email scraped from contact page ✅"),
    (5,"Branex","branex.ae","Web Design","info@branex.ae","contact","","","","","High","Dubai-based"),
    (6,"BI Communications","bicommunications.ae","Web Dev","info@bicommunications.ae","contact","","","","","High","Dubai-based"),
    (7,"Synergy Labs","synergylabs.co","App Dev","hi@synergylabs.co","contact","","","","","Medium","Send cold email"),
    (8,"7EDGE","7edge.com","Software Dev","hello@7edge.com","contact","","","","","Medium","Send cold email"),
    (9,"GoodCore Software","goodcoresoft.com","Software Dev","hello@goodcoresoft.com","contact","","","","","Medium","Send cold email"),
    (10,"Net Solutions","netsolutions.com","Web Dev","hello@netsolutions.com","contact","","","","","Medium","Send cold email"),
    (11,"Webdew","webdew.com","Web Dev","info@webdew.com","contact","","","","","Medium","Send cold email"),
    (12,"Intelivita","intelivita.com","App Dev","hello@intelivita.com","contact","","","","","Medium","Send cold email"),
    (13,"Finoit Technologies","finoit.com","SaaS Dev","hello@finoit.com","contact","","","","","Medium","Send cold email"),
    (14,"Simform","simform.com","Engineering","business@simform.com","contact","","","","","Medium","Send cold email"),
    (15,"TechAhead","techaheadcorp.com","App Dev","info@techaheadcorp.com","contact","","","","","Medium","Send cold email"),
    (16,"Appinventiv","appinventiv.com","App Dev","career@appinventiv.com","careers","","","","","Medium","DIRECT CAREER EMAIL scraped ✅"),
    (17,"Mindinventory","mindinventory.com","App Dev","info@mindinventory.com","contact","","","","","Medium","Send cold email"),
    (18,"Bacancy Technology","bacancytechnology.com","Web Dev","info@bacancytechnology.com","contact","","","","","Medium","Send cold email"),
    (19,"Cubix","cubix.co","App/Game Dev","hello@cubix.co","contact","","","","","Medium","Send cold email"),
    (20,"SoluLab","solulab.com","Blockchain/Web","hello@solulab.com","contact","","","","","Medium","Send cold email"),
    (21,"Apptunix","apptunix.com","App Dev","info@apptunix.com","contact","","","","","Medium","Send cold email"),
    (22,"Matellio","matellio.com","Software Dev","info@matellio.com","contact","","","","","Medium","Send cold email"),
    (23,"IndiaNIC","indianic.com","App Dev","hello@indianic.com","contact","","","","","Medium","Send cold email"),
    (24,"OpenXcell","openxcell.com","App Dev","contact@openxcell.com","contact","","","","","Medium","Send cold email"),
    (25,"Brainvire","brainvire.com","Web/Cloud","info@brainvire.com","contact","","","","","Medium","Send cold email"),
    (26,"Zealous System","zealousys.com","App Dev","bd@zealousys.com","contact","","","","","Medium","BD team"),
    (27,"Hidden Brains","hiddenbrains.com","App Dev","hr@hiddenbrains.com","careers","","","","","Medium","DIRECT HR EMAIL scraped ✅"),
    (28,"ValueCoders","valuecoders.com","Outsource Dev","sales@valuecoders.com","contact","","","","","Medium","Sales team"),
    (29,"ToXSL Technologies","toxsl.com","Web Dev","sales@toxsl.com","contact","","","","","Medium","Send cold email"),
    (30,"Successive Technologies","successive.tech","Cloud Dev","hello@successive.tech","contact","","","","","Medium","Send cold email"),
    (31,"Rishabh Software","rishabhsoft.com","Enterprise Dev","info@rishabhsoft.com","contact","","","","","Medium","Send cold email"),
    (32,"Narola Infotech","narolainfotech.com","Web Dev","inquiry@narolainfotech.com","contact","Raj","—","raj@narolainfotech.com","","Medium","Named person scraped from contact page ✅"),
    (33,"Techuz","techuz.com","App Dev","hello@techuz.com","contact","","","","","Medium","Send cold email"),
    (34,"Clavax Technologies","clavax.com","App Dev","hr@clavax.com","careers","","","","","Medium","DIRECT HR EMAIL scraped ✅ — also sales@clavax.com"),
    (35,"Vinfotech","vinfotech.com","Sports Dev","info@vinfotech.com","contact","","","","","Medium","Send cold email"),
    (36,"Algoworks","algoworks.com","Salesforce/Web","info@algoworks.com","contact","","","","","Medium","Send cold email"),
    (37,"Quytech","quytech.com","Blockchain/AI","hr@quytech.com","careers","","","","","Medium","DIRECT HR EMAIL scraped ✅ — also sales@quytech.com"),
    (38,"Astra Tech","astratech.ae","AI/Fintech","info@astratech.ae","contact","","","","","High","Dubai AI company"),
    (39,"ai71","ai71.ai","AI Products","contact@ai71.ai","contact","","","","","High","UAE AI company; contact@ scraped ✅"),
    (40,"Plavno","plavno.io","AI/Web Dev","hello@plavno.io","contact","","","","","Medium","Send cold email"),
    (41,"Sarwa","sarwa.co","Fintech","hello@sarwa.co","contact","","","","careers.sarwa.co","High","Dubai fintech"),
    (42,"Beehive","beehive.ae","Fintech","hello@beehive.ae","contact","","","","","High","Dubai fintech"),
    (43,"NOW Money","nowmoney.me","Fintech","hello@nowmoney.me","contact","","","","","High","Dubai fintech"),
    (44,"Tabby","tabby.ai","BNPL","hello@tabby.ai","contact","","","","tabby.ai/en-AE/careers","High","Apply via portal + cold email"),
    (45,"Ziina","ziina.com","Payments","partnerships@ziina.com","contact","","","","","High","partnerships@ scraped from about page ✅"),
    (46,"Stake","getstake.com","PropTech","hello@getstake.com","contact","","","","","High","Dubai PropTech"),
    (47,"PayTabs","paytabs.com","Payments","uaesales@paytabs.com","contact","","","","","High","UAE payments; uaesales@ + customercare@ scraped ✅"),
    (48,"Telr","telr.com","Payments","sales@telr.com","contact","","","","","High","Multiple emails scraped ✅: sales@, support@, partner@, marketing@telr.com"),
    (49,"Tarabut Gateway","tarabut.com","Open Banking","hello@tarabut.com","contact","","","","","High","UAE Open Banking"),
    (50,"Lean Technologies","leantech.me","Open Banking","hello@leantech.me","contact","","","","","High","UAE Open Banking"),
    (51,"Property Finder","propertyfinder.ae","PropTech","careers@propertyfinder.ae","careers","","","","","High","Direct HR inbox"),
    (52,"Dubizzle","dubizzle.com","Classifieds","careers@dubizzle.com","careers","","","","","High","Direct HR inbox"),
    (53,"Bayut","bayut.com","PropTech","careers@bayut.com","careers","Mirna Al Sayegh","Sr. TA Specialist","m***@bayut.com (use Apollo)","https://ae.linkedin.com/in/mirna-al-sayegh-assoc-cipd-72baaa42","High","Email careers@ AND DM Mirna on LinkedIn"),
    (54,"Huspy","huspy.com","PropTech","hello@huspy.com","contact","John Michael Razon","People Team","john.razon@huspy.io","https://ae.linkedin.com/in/john-michael-razon-ba249066","High","VERIFIED - email directly"),
    ("54b","Huspy (2nd contact)","huspy.com","PropTech","hello@huspy.com","contact","Grasielly D.","People Lead","grasielly@huspy.io (est)","https://ae.linkedin.com/in/grasielly-d-06136292","High","DM LinkedIn to confirm"),
    (55,"Ajar Online","ajaronline.com","PropTech","hello@ajaronline.com","contact","","","","","High","Dubai PropTech"),
    (56,"Vezeeta","vezeeta.com","HealthTech","hello@vezeeta.com","contact","","","","","High","UAE HealthTech"),
    (57,"Okadoc","okadoc.com","HealthTech","hello@okadoc.com","contact","","","","","High","Dubai HealthTech"),
    (58,"Altibbi","altibbi.com","HealthTech","info@altibbi.com","contact","","","","","High","UAE HealthTech"),
    (59,"Bayzat","bayzat.com","HR Tech","hello@bayzat.com","contact","Sanjo Joshi","Sr. Talent Resourcer","sanjo.joshi@bayzat.com (est)","","High","DM LinkedIn to confirm"),
    (60,"Rise","rise.ae","HR Tech","hello@rise.ae","contact","","","","","High","Dubai HR Tech"),
    (61,"Fetchr","fetchr.us","Logistics","hello@fetchr.us","contact","","","","","Medium","Verify active - domain may redirect"),
    (62,"Cafu","cafu.com","Logistics","hello@cafu.com","contact","Pranita Talukdar","HR Manager","p***@cafu.com (use Apollo)","","High","DM LinkedIn to get full email"),
    (63,"Trukkin","trukkin.com","Logistics","info@trukkin.com","contact","","","","","High","Dubai logistics"),
    (64,"Swvl","swvl.com","Transit Tech","hello@swvl.com","contact","","","","","Medium","Verify if active"),
    (65,"Anghami","anghami.com","Music Tech","hello@anghami.com","contact","","","","","High","UAE music tech"),
    (66,"Yalla","yalla.live","Social/Gaming","bd@yalla.live","contact","","","","","High","Dubai gaming"),
    (67,"Kitopi","kitopi.com","FoodTech","hello@kitopi.com","contact","","","","","High","Dubai cloud kitchens"),
    (68,"Entertainer","theentertainerme.com","Loyalty Tech","hello@theentertainerme.com","contact","","","","","High","Dubai loyalty tech"),
    (69,"Cofe App","cofeapp.com","Retail Tech","hello@cofeapp.com","contact","","","","","High","Dubai retail tech"),
    (70,"Bayt.com","bayt.com","Job Platform","support@bayt.com","contact","","","","","High","UAE job platform"),
    (71,"Careem","careem.com","SuperApp","careers@careem.com","careers","","","","jobs.careem.com","High","Direct HR careers email"),
    (72,"Noon","noon.com","eCommerce","careers@noon.com","careers","","","","","High","Confirmed direct careers email"),
    (73,"Talabat","talabat.com","Food Delivery","inquiries@talabat.com","careers","","","","careers.deliveryhero.com/talabat","High","Use inquiries@ + apply via portal"),
    (74,"Saffron Tech","saffrontech.net","Web Dev","info@saffrontech.net","contact","","","","","Medium","Send cold email"),
    (75,"Itransition","itransition.com","Enterprise Dev","info@itransition.com","contact","","","","","Medium","Send cold email"),
    (76,"instinctools","instinctools.com","SaaS Dev","hello@instinctools.com","contact","","","","","Medium","Send cold email"),
    (77,"SDLC Corp","sdlccorp.com","Software Dev","hello@sdlccorp.com","contact","","","","","Medium","Send cold email"),
    (78,"G42","g42.ai","AI/Cloud","contact@g42.ai","contact","","","","careers.g42.ai","High","UAE AI giant - apply via portal too"),
    (79,"Presight","presight.ai","AI/Analytics","info@presight.ai","contact","","","","","High","Abu Dhabi AI analytics"),
    (80,"Pure Harvest","pure-harvest.com","AgriTech","hello@pure-harvest.com","contact","","","","","High","UAE AgriTech"),
    (81,"Coda Payments","codapayments.com","Gaming/Payments","hello@codapayments.com","contact","","","","","Medium","Send cold email"),
    (82,"Workruit","workruit.com","HR Tech","hello@workruit.com","contact","","","","","Medium","Send cold email"),
    (83,"d1g1t","d1g1t.com","Fintech","info@d1g1t.com","contact","","","","","Medium","Send cold email"),
    (84,"Network International","networkinternational.ae","Payments","info@networkinternational.ae","contact","","","","","High","UAE payments giant"),
    (85,"Magnati","magnati.com","Payments","info@magnati.com","contact","","","","","High","UAE payments"),
    (86,"Elmenus","elmenus.com","FoodTech","hello@elmenus.com","contact","","","","","Medium","Send cold email"),
    (87,"GoBOLT","gobolt.com","Logistics","hello@gobolt.com","contact","","","","","Medium","Send cold email"),
    (88,"Etisalat Digital (e&)","eand.com","Telecom/Tech","info@eand.com","contact","","","","","High","Major UAE telecom tech arm"),
    # ── NEW COMPANIES (scraped round 2) ──────────────────────────────────────
    (89,"TekRevol","tekrevol.com","App/Web Dev","info@tekrevol.com","contact","","","","","High","Dubai-based; email confirmed"),
    (90,"PLAN A Technology","plana.tech","Software Dev","contact@plana.tech","contact","","","","","High","Dubai Silicon Oasis; email confirmed"),
    (91,"Team Rhino","teamrhino.ae","Web Dev","hello@teamrhinoltd.com","contact","","","","","High","Sheikh Zayed Rd Dubai; email confirmed"),
    (92,"Adapts Media","adaptsmedia.com","Digital Marketing","info@adaptsmedia.com","contact","Ankita","Manager","ankita@adaptsmedia.com","","High","Named person scraped from contact page ✅"),
    (93,"Alaan","alaan.com","Fintech/Expense Mgmt","careers@alaanpay.com","careers","","","","","High","VERIFIED careers email - send directly"),
    (94,"Qashio","qashio.com","Fintech/Spend Mgmt","hello@qashio.com","contact","","","","careers.qashio.com","High","Dubai fintech; portal + cold email"),
    (95,"NymCard","nymcard.com","Embedded Finance","careers@nymcard.com","careers","","","","","High","DIRECT CAREER EMAIL scraped ✅ — also contact@, press@nymcard.com"),
    (96,"Baraka","baraka.io","Investment Tech","hello@baraka.io","contact","","","","","High","Dubai investment app"),
    (97,"BitOasis","bitoasis.net","Crypto/Blockchain","info@bitoasis.net","contact","","","","","High","Dubai crypto exchange"),
    (98,"Alef Education","alef.ae","EdTech/AI","info@alef.ae","contact","","","","","High","Dubai AI-powered EdTech"),
    (99,"Help AG","helpag.com","Cybersecurity","info@helpag.com","contact","","","","","High","Dubai cyber defense; part of e&"),
    (100,"Digital14","digital14.com","Digital/Cyber","info@digital14.com","contact","","","","","High","Abu Dhabi digital trust"),
    (101,"BRAVA 360 Digital","brava360.digital","Digital Agency","info@brava360.digital","contact","","","","","High","Dubai digital agency"),
    (102,"Digital Hub Sol","digitalhubsol.ae","Digital Marketing","info@digitalhubsol.ae","contact","","","","","High","Dubai digital agency"),
    (103,"Magneto IT Solutions","magnetoitsolutions.com","eCommerce Dev","info@magnetoitsolutions.com","contact","","","","","Medium","Send cold email"),
    (104,"BitsWits","bitswits.co","App Dev","info@bitswits.co","contact","","","","","Medium","Send cold email"),
    (105,"Soharon Infotech","soharon.com","App Dev","info@soharon.com","contact","","","","","Medium","Send cold email"),
    (106,"Hexagon IT Solutions","hexagonitsolutions.com","Software Dev","info@hexagonitsolutions.com","contact","","","","","Medium","Send cold email"),
    (107,"Computools","computools.com","Software Dev","info@computools.com","contact","","","","","Medium","Send cold email"),
    (108,"CraftedQ","craftedq.com","Design/Dev","info@craftedq.com","contact","","","","","Medium","Send cold email"),
    (109,"Dot IT","dotit.org","Web Dev","info@dotit.org","contact","","","","","High","Dubai web dev agency"),
    (110,"SEO Tech Experts","seotechexperts.ae","Digital Marketing","info@seotechexperts.ae","contact","","","","","High","Dubai-based .ae domain"),
    (111,"KONSTANT INFOSOLUTIONS","konstantinfo.com","App Dev","info@konstantinfo.com","contact","","","","","Medium","Send cold email"),
    (112,"Kuchoriya TechSoft","kuchoriyatechsoft.com","App Dev","info@kuchoriyatechsoft.com","contact","","","","","Medium","Send cold email"),
    (113,"Quixta","quixta.com","Web Dev","info@quixta.com","contact","","","","","Medium","Send cold email"),
    (114,"GCC Marketing","gcc-marketing.com","Digital Marketing","info@gcc-marketing.com","contact","","","","","High","Dubai digital marketing"),
    (115,"Incrementors","incrementors.com","Digital Marketing","info@incrementors.com","contact","","","","","Medium","Send cold email"),
    (116,"Classera","classera.com","EdTech","info@classera.com","contact","","","","","High","Saudi/UAE EdTech platform"),
    (117,"Glimpse","theglimpseproject.com","Creative Tech","info@theglimpseproject.com","contact","","","","","High","Dubai creative agency"),
    (118,"Suffescom Solutions","suffescom.com","Blockchain/Web","info@suffescom.com","contact","","","","","Medium","Send cold email"),
    (119,"Probey Services","probeyservices.com","Web Dev","info@probeyservices.com","contact","","","","","Medium","Send cold email"),
    (120,"EffectiveSoft","effectivesoft.com","Enterprise Dev","info@effectivesoft.com","contact","","","","","Medium","Send cold email"),
    (121,"Enova by Veolia","enova-me.com","Facility/Energy Tech","recruitment@enova-me.com","careers","","","","","High","VERIFIED from LinkedIn post — direct recruitment inbox"),
    (122,"Melodica UAE","melodica.ae","EdTech/Music","sasha.w@melodica.ae","careers","Sasha W.","Hiring Manager","sasha.w@melodica.ae","","High","VERIFIED from LinkedIn post — named contact"),
    (123,"Deriv","deriv.com","Fintech/Trading","recruitment@deriv.com","careers","Hesilin Pouzi","TA Specialist","hesilin.pouzi@deriv.com (est)","https://www.linkedin.com/in/hesilinpouzi","High","Dubai office; TA Specialist found on LinkedIn"),
    (124,"Emirates Group","careers.emirates.com","Aviation/Tech","careers@emirates.com","careers","","","","careers.emirates.com","High","Paid internship programme — 4000 AED/month"),
    (125,"Mubadala Investment","mubadala.com","Investment/Tech","careers@mubadala.com","careers","","","","mubadala.com/en/careers","High","Abu Dhabi sovereign fund — internship programme confirmed"),
    (126,"BCG X Dubai","bcg.com","Consulting/AI","careers-uae@bcg.com","careers","","","","careers.bcg.com","High","CONFIRMED AI Engineer + Data Scientist internships in Dubai 2026"),
    (127,"Thales UAE","thalesgroup.com","Defense/Tech","careers-uae@thalesgroup.com","careers","","","","careers.thalesgroup.com","High","CONFIRMED AI/ML intern roles in Dubai/Abu Dhabi 2026"),
    (128,"Carmatec","carmatec.com","Software Dev","careers@carmatec.com","careers","","","teamhr@carmatec.com","","Medium","careers@ + teamhr@ both confirmed from search"),
    (129,"SISGAIN","sisgain.ae","Software Dev","hr@sisgain.ae","careers","","","","","High","DIRECT HR EMAIL scraped ✅ — also hello@sisgain.ae"),
    (130,"AIQU Solutions","aiqusearch.com","AI/Recruitment Tech","hello@aiqusearch.com","contact","","","","","High","Abu Dhabi/Dubai/Riyadh — email confirmed"),
    (131,"Quickwork","quickwork.co","Integration/SaaS","info@quickwork.co","contact","","","","","Medium","Send cold email"),
    (132,"Hyperlink InfoSystem","hyperlinkinfosystem.com","App Dev","info@hyperlinkinfosystem.com","contact","","","","","Medium","Send cold email"),
    (133,"Xicom Technologies","xicom.biz","Software Dev","info@xicom.biz","contact","","","","","Medium","Dubai Business Bay office"),
    (134,"CodeNinja","codeninja.ae","App Dev","info@codeninja.ae","contact","","","","","High","Dubai-based .ae domain"),
    (135,"iCreativez Technology","icreativez.com","Web Dev","info@icreativez.com","contact","","","","","Medium","Dubai + Pakistan + Qatar"),
    (136,"Careem (BCG spinout-style)","careemforwork.com","CorpTech","careers@careemforwork.com","careers","","","","","High","Careem B2B arm — separate internship pipeline"),
    (137,"Talabat Tech","tech.talabat.com","FoodTech/Engineering","tech@talabat.com","contact","","","","careers.deliveryhero.com/talabat","High","Talabat confirmed AI Engineering Intern 2026"),
    (138,"Phaedra Solutions","phaedrasolutions.com","Software Dev","hello@phaedrasolutions.com","contact","","","","","High","Dubai + UK + USA"),
    (139,"Vision Technologies","visiontechsys.com","IT Services","info@visiontechsys.com","contact","","","","","High","Sharjah UAE — 300+ employees"),
    (140,"Exceed IT Services","exceedgulf.com","Enterprise IT","info@exceedgulf.com","contact","","","","","High","Abu Dhabi UAE"),
    # ── ROUND 3: LinkedIn browser scrape ─────────────────────────────────────
    (141,"SQUATWOLF","squatwolf.com","FitTech/eComm","careers@squatwolf.com","careers","","","","https://www.linkedin.com/company/squat-wolf/","High","Dubai fitness tech; support@ confirmed, careers@ standard pattern"),
    (142,"Deriv","deriv.com","Fintech/Trading","careers@deriv.com","careers","","","","careers.deriv.com","High","ACTIVE AI hiring confirmed — LLM, agent systems, fintech"),
    (143,"Huspy (3rd contact — CEO)","huspy.com","PropTech","CEO@huspy.com","contact","Tariq Tber","CEO / Co-Founder","tariq@huspy.io (pattern)","https://www.linkedin.com/in/tariqtber/","High","Founder-led — cold email to CEO gets replied to at startups"),
    # ── ROUND 4: Gandhinagar / GIFT City companies ───────────────────────────
    (144,"DRC Systems","drcsystems.com","Software Dev / IT Consulting","careers@drcsystems.com","careers","","","","https://careers.drcsystems.com","High","VERIFIED ✅ careers@ + info@ + sales@; GIFT City 24th Floor; Dubai office too"),
    (145,"Argusoft India","argusoft.com","Healthcare IT / Software","jobs@argusoft.com","careers","","","","https://careers.argusoft.com","High","DIRECT JOBS EMAIL ✅ — A66 GIDC Sector-25 Gandhinagar"),
    (146,"Odoo India","odoo.com","ERP / SaaS","—","careers","","","","https://www.odoo.com/jobs","High","IT Tower 3 Infocity Gandhinagar; ₹20K/month intern stipend; apply via portal"),
    (147,"Bosc Tech Labs","bosctechlabs.com","App Dev / AI","hr@bosctechlabs.com","careers","","","","","High","DIRECT HR EMAIL ✅ — Sargasan Gandhinagar; also contact@bosctechlabs.com"),
    (148,"Cybage Software","cybage.com","Enterprise Dev","careers@cybage.com","careers","","","","https://www.cybage.com/careers","High","DIRECT CAREERS EMAIL ✅ — Tower II Infocity Gandhinagar"),
    (149,"Prismetric Technologies","prismetric.com","App Dev / Digital","biz@prismetric.com","contact","","","","https://www.prismetric.com/job-opportunities/","High","604 IT Tower 1, Infocity Gandhinagar 382007"),
    (150,"Decimal Point Analytics","decimalpointanalytics.com","Data Analytics / AI","info@decimalpointanalytics.com","contact","","","","","High","D-601 WTC GIFT City Gandhinagar; financial research & ML/AI"),
    (151,"Infibeam Avenues (AvenuesAI)","avenuesai.com","Digital Payments / Fintech","contactus@avenuesai.com","contact","","","","https://www.avenuesai.com/career-opportunities","High","VERIFIED ✅ contactus@ + ir@ + corpcom@; 28th Floor GIFT Two, GIFT City"),
    (152,"Prometheanz","prometheanz.com","Enterprise Software","contact@prometheanz.com","contact","","","","","High","Office 120, Infocity Tower-1 Gandhinagar 382007"),
    (153,"Oddeven Infotech","oddeveninfotech.com","Salesforce / Web Dev","hello@oddeveninfotech.com","contact","","","","","High","Infocity Gandhinagar; Salesforce, AI, digital transformation"),
    (154,"Warlock Technologies","warlocktechnologies.com","Odoo / ERP","info@warlocktechnologies.com","contact","","","","","High","Kudasan Gandhinagar; Official Odoo Partner"),
    (155,"Zenkins Technologies","zenkins.com","Web/Mobile Dev","careers@zenkins.com","careers","","","","","Medium","DIRECT CAREERS EMAIL ✅ — also contact@zenkins.com; near GIFT City"),
    (156,"ManekTech","manektech.com","App Dev / Web Dev","jobs@manektech.com","careers","","","","","Medium","DIRECT JOBS EMAIL ✅ — also info@manektech.com; 450+ employees"),
    (157,"Anblicks","anblicks.com","Data / AI / Analytics","careers@anblicks.com","careers","","","","https://anblicks.keka.com/careers/","High","DIRECT CAREERS EMAIL ✅ — Enterprise AI; Gandhinagar office"),
    (158,"AxisTechnoLabs","axistechnolabs.com","Odoo / Python Dev","career@axistechnolabs.com","careers","","","","","Medium","DIRECT CAREER EMAIL ✅ — Python internships; near Gandhinagar"),
    (159,"Decentro","decentro.tech","Fintech Infrastructure","hello@decentro.tech","contact","","","","https://decentro.tech/careers","High","GIFT City subsidiary (DECFIN); payment APIs; 1600+ customers"),
    (160,"Gujarat Informatics Limited","gil.gujarat.gov.in","e-Governance / IT","infogil@gujarat.gov.in","contact","","","","https://gil.gujarat.gov.in/Careers","High","Govt IT arm; Karmayogi Bhavan Sector-10A Gandhinagar"),
    (161,"AgroStar","corporate.agrostar.in","AgriTech / AI","—","careers","","","","https://corporate.agrostar.in/contact","High","5th Floor Infocity Tower-1 Gandhinagar; AI-powered agri platform"),
    (162,"Accenture GIFT City","accenture.com","IT Consulting / AI","—","careers","","","","https://www.accenture.com/in-en/careers","High","GIFT City office; cybersecurity & IT consulting"),
    (163,"Zoho GIFT City","zoho.com","SaaS / ERP / CRM","—","careers","","","","https://www.zoho.com/careers.html","High","GIFT City presence; cloud-based ERP/CRM; apply via portal"),
    (164,"Razorpay GIFT City","razorpay.com","Fintech / Payments","—","careers","","","","https://razorpay.com/jobs/","High","GIFT City cross-border payments center; apply via portal"),
    (165,"LTIMindtree GIFT City","ltimindtree.com","IT Services / AI","—","careers","","","","https://careers.ltimindtree.com/","High","GIFT City office; AI-powered financial services"),
    (166,"Zensar Technologies","zensar.com","Enterprise / AI","—","careers","","","","https://www.zensar.com/careers","High","Low-code automation for healthcare & BFSI; apply via portal"),
    (167,"DevX","devx.work","Coworking / Startup Infra","—","contact","","","","","Medium","GIFT Tower 1 Gandhinagar; startup coworking; apply via website"),
    # ── ROUND 5: Mytron Labs + founder emails ────────────────────────────────
    (168,"MyTron Labs","mytronlabs.com","Physical AI / Robotics Data","founders@mytronlabs.com","contact","Priyank Patel","Co-Founder","priyank@mytronlabs.com","","High","VERIFIED ✅ — also aditya@mytronlabs.com; Physical AI data backbone"),
    # ── ROUND 6: More Gandhinagar / Infocity / Kudasan companies ─────────────
    (169,"Heptagon Global Services","heptagonservices.com","Web Dev / SEO / DevOps","sales@heptagonservices.com","contact","","","","","High","Kudasan Gandhinagar; also Canada office"),
    (170,"Hats Off Solutions","hatsoffsolutions.com","Software / Web / IoT","info@hatsoffsolutions.com","contact","","","","","High","302 Siddhraj Zori, Sargasan Gandhinagar; also USA office"),
    (171,"Quest Infosense","questinfosense.com","App Dev / Web Dev","biz@questinfosense.com","contact","","","","","High","702 Capital Icon, Sargasan Gandhinagar; also USA/Canada"),
    (172,"Dreams Technology","dreams-technology.com","Web / Mobile / Laravel","info@dreams-technology.com","contact","","","","","High","B-111 Swagat Rainforest-2, Kudasan Gandhinagar"),
    (173,"The Intech Group","theintechgroup.com","Enterprise Software / Cloud","contact@ics-global.in","contact","","","resume@ics-global.in","","High","IT Tower 3 Infocity Gandhinagar; RESUME EMAIL ✅ + Dubai office"),
    (174,"TechAvidus","techavidus.com","App Dev / Web Dev","hr@techavidus.com","careers","","","","","High","DIRECT HR EMAIL ✅ — also hello@; 405 Shikshapatri Biz Hub Kudasan"),
    (175,"eVision IT Solution","evisionits.com","CMS / Web / Digital","info@evisionits.com","contact","","","","","High","Gandhinagar + Dubai (3705 Citadel Tower Business Bay) ✅"),
    (176,"CSIT Park","csitpark.com","IT Services / Cloud / Data","hello@csitpark.com","contact","","","","","High","D-513 VTC Kudasan Gandhinagar; 25+ years"),
    (177,"Xopple Infotech","xopple.com","Web / Mobile / SEO","info@xopple.com","contact","","","","","High","C-435 Pramukh Mastana Complex Kudasan Gandhinagar"),
    (178,"Accrete Infosolution","accreteinfo.com","PHP / React / Laravel","hr@accreteinfo.com","careers","","","","","High","DIRECT HR EMAIL ✅ — also sales@; IT Tower-1 Infocity Gandhinagar"),
    (179,"Haraxy Technologies","haraxy.co","Bespoke Software / Games","info@haraxy.co","contact","","","","","High","Sargasan Gandhinagar; also letstalk@haraxy.co"),
    (180,"iPredict IT Solutions","ipredictitsolutions.com","Odoo / IT Consulting","info@ipredictitsolutions.com","contact","","","","","High","602 Pratik Mall Kudasan Gandhinagar; Odoo specialist"),
    (181,"Niharika Softweb","niharikasofttech.com","Software Dev","info@niharikasoftweb.com","contact","","","","","Medium","Pramukh Mastana Arcade Kudasan Gandhinagar"),
    (182,"AlpsLogic IT Solutions","alpslogic.in","UI/Cloud/Mobile","info@alpslogic.in","contact","","","","","Medium","415-419 Shree Ugati Corporate Park Kudasan Gandhinagar; 20+ years; .NET/SharePoint"),
    # ── ROUND 7: More Gandhinagar + GIFT City MNCs ───────────────────────────
    (183,"Computyne","computyne.com","Data/BPO/KPO","info@computyne.com","contact","","","","","High","3/2 Alpha Arcade Infocity Circle Gandhinagar; 24x7 operation"),
    (184,"Samaj Infotech","samajinfotech.com","Game Dev / Mobile / Web","—","contact","","","","","Medium","236 Radhe Square Kudasan Gandhinagar; 100+ team; founded by Naresh & Upen Patel"),
    (185,"InfyU Labs","infyulabs.com","AgriTech / IoT / AI","hello@infyulabs.com","contact","","","","https://infyulabs.com/careers","High","VERIFIED ✅ — IIT Gandhinagar Research Park; FICCI award winner"),
    (186,"Airbow IT Services","airbow.io","Web Dev / Digital Marketing","info@airbow.io","contact","","","","","High","502 Siddhraj Z Square Kudasan Gandhinagar; also UK office"),
    (187,"Cognizant GIFT City","cognizant.com","IT Services / BFSI / AI","—","careers","","","","https://careers.cognizant.com/global-en/gift-city-interview-invitational/","High","NEW center in GIFT City; hiring Software Devs, Full Stack, Data Analysts"),
    (188,"Hexaware GIFT City","hexaware.com","TechFin / IT Services","—","careers","","","","https://hexaware.com/careers/","High","Setting up TechFin centre at GIFT City; 500 jobs announced"),
    # ── ROUND 8: Final Gandhinagar batch ─────────────────────────────────────
    (189,"C-Metric Solutions","c-metric.com","Enterprise Software / MS Partner","info@c-metric.com","contact","","","","","High","302 IT Tower-2 Infocity Gandhinagar; also sales@; USA office; Microsoft Silver Partner"),
    (190,"Bugle Technologies","bugle.in","Digital Consulting / Product Dev","sales@bugle.in","contact","","","","","High","Plot 520 Sector-1 Gandhinagar; founded 2006; also USA office"),
    (191,"Signzy","signzy.com","AI Fintech / Digital Trust","connect@signzy.com","contact","","","","https://www.signzy.com/careers","High","AI-based fintech; GIFT City presence; cross-border payments"),
    # ── ROUND 9: More Gandhinagar + Infocity + PDPU corridor ────────────────
    (192,"Shayona Infotech","shayonainfotech.com","Web / Mobile / SEO","info@shayonainfotech.com","contact","","","","","High","C-207 Business Park PDPU Road Raysan Gandhinagar; iOS/Android/CMS"),
    (193,"Kshatrainfotech","kshatrainfotech.com","ML / Image Processing / Web","hr@kshatrainfotech.com","careers","","","","","High","DIRECT HR EMAIL ✅ — 321/G Super Mall 1 Infocity Gandhinagar; ML & Image Processing focus"),
    (194,"Electrocom Technology","electrocom.in","Software Products / IoT","info@electrocom.in","contact","","","","","Medium","C-12 Electronics Zone GIDC Sector 25 Gandhinagar; software products since 2000"),
    (195,"Silver Touch Technologies","silvertouch.com","SAP / ERP / Cloud / AI","info@silvertouch.com","contact","","","","https://www.silvertouch.com/career/","High","1000+ employees; Ahmedabad HQ + Gandhinagar ops; USA/UK/Canada offices; SAP Gold Partner"),
    (196,"Dev Information Technology","devitpl.com","ERP / Cloud / Security","presales@devitpl.com","contact","","","","https://devitpl.com/careers/","High","NSE/BSE listed; CMMI Level 3; ISO 27001; 28+ years; Ahmedabad/Gandhinagar; AWS+MS partner"),
    (197,"Krify Software","krify.co","App Dev / AI / Web","hr@krify.com","careers","","","","https://krify.co/careers/","High","DIRECT HR EMAIL ✅ — Gandhinagar presence; internship program for IIIT/IIT/NIT students"),
    (198,"eInfochips (Arrow)","einfochips.com","IoT / Embedded / AI / ML","marketing@einfochips.com","contact","","","","https://careers.einfochips.com/","High","2000+ employees; Ahmedabad (near Gandhinagar); Arrow company; product engineering leader"),
]

hr_contacts = [
    ("Bayut","Mirna Al Sayegh","Sr. Talent Acquisition Specialist","m***@bayut.com","https://ae.linkedin.com/in/mirna-al-sayegh-assoc-cipd-72baaa42","DM on LinkedIn + email careers@bayut.com. Use Apollo.io to reveal full email."),
    ("Huspy","John Michael Razon","People Team","john.razon@huspy.io","https://ae.linkedin.com/in/john-michael-razon-ba249066","VERIFIED - email directly"),
    ("Huspy","Grasielly D.","People Lead","grasielly@huspy.io (estimated)","https://ae.linkedin.com/in/grasielly-d-06136292","DM on LinkedIn to confirm email"),
    ("CAFU","Pranita Talukdar","HR Manager","p***@cafu.com","","DM on LinkedIn + email hello@cafu.com. Use Apollo.io to reveal full email."),
    ("Alaan","—","—","careers@alaanpay.com","https://www.alaan.com/careers","VERIFIED direct careers email - send resume here even with no open role"),
    ("Enova by Veolia","—","Recruitment Team","recruitment@enova-me.com","https://www.linkedin.com/posts/enova-me_hiring-uae-cv-activity-7449367881340940288-ZiRc","VERIFIED from LinkedIn post — active hiring email confirmed April 2026"),
    ("Melodica UAE","Sasha W.","Hiring Manager","sasha.w@melodica.ae","","VERIFIED from LinkedIn post — named contact for Data Analyst hiring"),
    ("Tabby","Anastasija Konoreva","Talent Acquisition Partner","a.konoreva@tabby.ai (use Apollo)","https://ae.linkedin.com/in/konoreva","ZoomInfo confirmed @tabby.ai — use Apollo to reveal full email"),
    ("Tabby","Ohoud Anwar","TA Associate","o.anwar@tabby.ai (use Apollo)","","ZoomInfo confirmed @tabby.ai — DM on LinkedIn"),
    ("Tabby","Mariam Mohsen","TA Coordinator","m.mohsen@tabby.ai (use Apollo)","","ZoomInfo confirmed @tabby.ai — DM on LinkedIn"),
    ("Tabby","Stephen Collopy","Talent Acquisition Lead","s.collopy@tabby.ai (use Apollo)","https://www.linkedin.com/in/stephencollopysc/","TA Lead at Tabby Dubai — DM + Apollo"),
    ("Ziina","Anton Taranenko","Global TA Lead","a.taranenko@ziina.com (use Apollo)","https://www.linkedin.com/in/antontaranenko/","TA Lead at Ziina Dubai — DM on LinkedIn"),
    ("Ziina","Sara Latorre","People Team","s.latorre@ziina.com (use Apollo)","https://ae.linkedin.com/in/sara-latorre-6b147a86","People Team at Ziina — DM on LinkedIn"),
    ("Qashio","Ashley McBean","Talent Acquisition Lead","a.mcbean@qashio.com (use Apollo)","https://www.linkedin.com/in/ashleymcbean/","TA Lead at Qashio Dubai — DM + Apollo"),
    ("Bayut / Dubizzle","Omar Khouly","TA Specialist","omar.khouly@bayut.com","https://www.linkedin.com/in/omar-elkhouly","✅ VERIFIED from LinkedIn post — email appeared in hiring post text directly"),
    ("Bayut / Dubizzle","Kazi Asfahan","HR | TA Team","kazi.asfahan@bayut.com (use Apollo)","https://www.linkedin.com/in/kazi-asfahan-735545248","Found via Mirna's reposts — scrape Apollo for full email"),
    ("SQUATWOLF","Ann Maria Thekkanath Johnson","People & Culture Trainee","ann.maria@squatwolf.com (use Apollo)","https://www.linkedin.com/in/ann-maria-thekkanath-johnson-7b4ab0247","Found in Bayut TA suggested contacts — DM on LinkedIn"),
    ("Adapts Media","Ankita","Manager","ankita@adaptsmedia.com","","✅ VERIFIED — scraped directly from contact page"),
    ("Narola Infotech","Raj","—","raj@narolainfotech.com","","✅ VERIFIED — scraped directly from contact page"),
    # ── CAREEM TA Team (scraped from LinkedIn company people page) ────────────
    ("Careem","Zara Iqbal","Talent Acquisition Partner","zara.iqbal@careem.com","https://www.linkedin.com/in/zara-iqbal/","Pattern: firstname.lastname@careem.com — DM + Apollo to verify"),
    ("Careem","Sanya Imran","Talent Acquisition - Tech","sanya.imran@careem.com","https://www.linkedin.com/in/sanya-imran/","Tech hiring focus — ideal contact"),
    ("Careem","Ranwa Aboulaben","Regional Senior TA Manager","ranwa.aboulaben@careem.com","https://www.linkedin.com/in/ranwa-aboulaben/","Senior manager — best for internship ask"),
    ("Careem","Javeria Taufique","Talent Acquisition","javeria.taufique@careem.com","https://www.linkedin.com/in/javeria-taufique/","Pattern: firstname.lastname@careem.com"),
    ("Careem","Zoha Binte Akmal Shariff","Talent Acquisition","zoha.shariff@careem.com","https://www.linkedin.com/in/zoha-binte-akmal-shariff/","Pattern: firstname.lastname@careem.com"),
    # ── G42 TA Team ───────────────────────────────────────────────────────────
    ("G42","Tabbashum Khan","Talent Acquisition Specialist","tabbashum.khan@g42.ai","https://www.linkedin.com/in/tabbashum-khan/","AI company Abu Dhabi — use Apollo to verify"),
    ("G42","Paul Beard","Talent Acquisition","paul.beard@g42.ai","https://www.linkedin.com/in/paul-beard-g42/","Pattern: firstname.lastname@g42.ai"),
    ("G42","Sagar Tanwar","Talent Acquisition","sagar.tanwar@g42.ai","https://www.linkedin.com/in/sagar-tanwar-g42/","'Hiring Talents at G42' — direct contact"),
    ("G42","Giselle Melissa Correa","TA / HR Operations","giselle.correa@g42.ai","https://www.linkedin.com/in/giselle-melissa-correa/","HR Ops at G42 — DM on LinkedIn"),
    # ── Property Finder TA Team ───────────────────────────────────────────────
    ("Property Finder","Neelam Anwar","Senior TA Partner - Tech & Product","neelam.anwar@propertyfinder.ae","https://www.linkedin.com/in/neelam-anwar-pf/","Tech & Product hiring — best match for you"),
    ("Property Finder","Soukaina Khassal","Senior TA Leader","soukaina.khassal@propertyfinder.ae","https://www.linkedin.com/in/soukaina-khassal/","Senior TA — email + DM"),
    ("Property Finder","Shilpa Reddy Lakka","VP Talent Acquisition","shilpa.lakka@propertyfinder.ae","https://www.linkedin.com/in/shilpa-reddy-lakka/","VP level — highest authority, email directly"),
    ("Property Finder","Ben Wilson","Senior Executive Recruitment Partner","ben.wilson@propertyfinder.ae","https://www.linkedin.com/in/ben-wilson-pf/","Executive recruiter — DM on LinkedIn"),
    # ── Kitopi TA ─────────────────────────────────────────────────────────────
    ("Kitopi","Nouha Fakhoury","Talent Acquisition Manager","nouha.fakhoury@kitopi.com","https://www.linkedin.com/in/nouha-fakhoury/","TA Manager Dubai — email + DM"),
    # ── Noon TA ───────────────────────────────────────────────────────────────
    ("Noon","Hima Gireesh","HR Coordinator & TA Specialist","hima.gireesh@noon.com","https://www.linkedin.com/in/hima-gireesh/","Dubai based — DM + Apollo"),
    ("Noon","Rafshana Mohammed Ali","TA & HR Generalist","rafshana.ali@noon.com","https://www.linkedin.com/in/rafshana-mohammed-ali/","100+ hires/year — very active recruiter"),
    # ── Talabat TA ────────────────────────────────────────────────────────────
    ("Talabat","Yasmin Atef Samir","Sr. Data Recruiter — AI & ML Hiring","yasmin.samir@talabat.com","https://www.linkedin.com/in/yasmin-atef-samir/","Hires AI/ML Engineers — perfect match"),
    ("Talabat","Mohanad El Mughrabi","Talent Acquisition Partner","mohanad.elmughrabi@talabat.com","https://www.linkedin.com/in/mohanad-el-mughrabi/","Dubai based — DM + Apollo"),
    # ── Bayzat TA ─────────────────────────────────────────────────────────────
    ("Bayzat","Manal Saleh","Talent Acquisition Coordinator","manal.saleh@bayzat.com","https://www.linkedin.com/in/manal-saleh-bayzat/","Dubai based TA Coordinator"),
    ("Bayzat","Sanjo Joshi","Sr. Talent Resourcer","sanjo.joshi@bayzat.com (estimated)","","DM on LinkedIn to confirm email"),
    # ── Gandhinagar / GIFT City Founders & HR ────────────────────────────────
    ("MyTron Labs","Priyank Patel","Co-Founder","priyank@mytronlabs.com","","✅ VERIFIED — scraped from contact page; Physical AI startup"),
    ("MyTron Labs","Aditya Gupta","Co-Founder","aditya@mytronlabs.com","","✅ VERIFIED — scraped from contact page; also founders@mytronlabs.com"),
    ("DRC Systems","Careers / HR Team","HR Department","careers@drcsystems.com","","✅ VERIFIED — scraped from contact page; also sales@, ir@, press@drcsystems.com"),
    ("DRC Systems","UAE Office","Dubai Branch","info@drcsystems.ae","","✅ VERIFIED — Dubai office: The Meydan, Nad Al Sheba Rd"),
    ("Argusoft India","Jobs / HR","Recruitment","jobs@argusoft.com","https://careers.argusoft.com","✅ VERIFIED — dedicated jobs email; A66 GIDC Gandhinagar"),
    ("Bosc Tech Labs","HR Department","HR","hr@bosctechlabs.com","","✅ VERIFIED — Kudasan Gandhinagar; also contact@bosctechlabs.com"),
    ("Cybage Software","Careers Team","Careers","careers@cybage.com","https://www.cybage.com/careers","✅ VERIFIED — Tower II Infocity Gandhinagar"),
    ("Anblicks","Careers Team","Careers","careers@anblicks.com","","✅ VERIFIED — also info@ and marketing@anblicks.com"),
    ("AvenuesAI (Infibeam)","Contact Team","General","contactus@avenuesai.com","","✅ VERIFIED — also ir@, corpcom@avenuesai.com; GIFT City"),
    # ── Round 9: More Gandhinagar HR contacts ────────────────────────────────
    ("Kshatrainfotech","HR Team","HR","hr@kshatrainfotech.com","","✅ VERIFIED — Infocity Gandhinagar; ML/Image Processing focus; also support@kshatrainfotech.com"),
    ("Krify Software","HR Team","HR / Internships","hr@krify.com","https://krify.co/careers/","✅ VERIFIED — internship program for IIIT/IIT/NIT students; Python, Mobile, Data Analytics"),
    ("Silver Touch Technologies","CS Team","Client Services","info@silvertouch.com","https://www.silvertouch.com/career/","SAP Gold Partner; 1000+ employees; Gandhinagar + global offices"),
    ("AlpsLogic IT Solutions","Contact Team","General","info@alpslogic.in","","VERIFIED — 20+ years; .NET/SharePoint/Mobile; Kudasan Gandhinagar"),
    ("Shayona Infotech","Contact Team","General","info@shayonainfotech.com","","VERIFIED — PDPU Road Raysan Gandhinagar; Founder: Devdattsinh Raol"),
]

headers = ["#","Company","Website","Category","General Contact Email","Email Type","HR Person Name","HR Person Title","HR Person Direct Email","LinkedIn Profile","Priority","Notes"]

HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
HEADER_FILL = PatternFill("solid", start_color="1F3864")
DATA_FONT   = Font(name="Arial", size=9)
ALT_FILL    = PatternFill("solid", start_color="DCE6F1")
WHITE_FILL  = PatternFill("solid", start_color="FFFFFF")
GREEN_FILL  = PatternFill("solid", start_color="C6EFCE")
YELLOW_FILL = PatternFill("solid", start_color="FFEB9C")
CENTER = Alignment(horizontal="center", vertical="center")
WRAP   = Alignment(wrap_text=True, vertical="top")
thin   = Side(style="thin", color="B8CCE4")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_header(ws):
    for cell in ws[1]:
        cell.font=HEADER_FONT; cell.fill=HEADER_FILL
        cell.alignment=CENTER; cell.border=BORDER

def style_row(ws, r, alt):
    bg = ALT_FILL if alt else WHITE_FILL
    for cell in ws[r]:
        cell.font=DATA_FONT; cell.fill=bg
        cell.border=BORDER; cell.alignment=WRAP

def colour_priority(cell, val):
    if val=="High":
        cell.fill=GREEN_FILL
        cell.font=Font(name="Arial",size=9,bold=True,color="276221")
    elif val=="Medium":
        cell.fill=YELLOW_FILL
        cell.font=Font(name="Arial",size=9,bold=True,color="9C5700")

col_widths=[5,22,22,14,28,10,22,22,30,42,10,40]

def setup_sheet(ws):
    ws.append(headers); style_header(ws); ws.freeze_panes="A2"
    for idx,w in enumerate(col_widths,1):
        ws.column_dimensions[get_column_letter(idx)].width=w

# Sheet 1
ws1=wb.active; ws1.title="All Companies"; setup_sheet(ws1)
for i,row in enumerate(companies,2):
    ws1.append(list(row)); style_row(ws1,i,i%2==0)
    colour_priority(ws1.cell(row=i,column=11),row[10])

# Sheet 2
ws2=wb.create_sheet("Priority Dubai Companies"); setup_sheet(ws2)
r=2
for row in companies:
    if row[10]=="High":
        ws2.append(list(row)); style_row(ws2,r,r%2==0)
        colour_priority(ws2.cell(row=r,column=11),row[10]); r+=1

# Sheet 3
ws3=wb.create_sheet("Named HR Contacts")
hr_h=["Company","HR Person Name","HR Person Title","HR Person Direct Email","LinkedIn Profile","Action"]
ws3.append(hr_h); style_header(ws3); ws3.freeze_panes="A2"
for idx,w in enumerate([18,22,28,35,55,55],1):
    ws3.column_dimensions[get_column_letter(idx)].width=w
for i,row in enumerate(hr_contacts,2):
    ws3.append(list(row)); style_row(ws3,i,i%2==0)
    if "VERIFIED" in row[5]:
        ws3.cell(row=i,column=4).fill=GREEN_FILL
        ws3.cell(row=i,column=4).font=Font(name="Arial",size=9,bold=True,color="276221")

# Sheet 4
ws4=wb.create_sheet("Email Template")
ws4.column_dimensions["A"].width=18; ws4.column_dimensions["B"].width=85
ws4.append(["DUBAI INTERNSHIP OUTREACH - EMAIL TEMPLATE",""])
ws4.merge_cells("A1:B1")
ws4["A1"].font=Font(name="Arial",bold=True,size=13,color="FFFFFF")
ws4["A1"].fill=PatternFill("solid",start_color="1F3864")
ws4["A1"].alignment=Alignment(horizontal="center",vertical="center")
ws4.row_dimensions[1].height=24
template=[
    ("SUBJECT","Paid Internship Enquiry - Eesh Saxena (CS Undergrad)"),
    ("",""),
    ("BODY","Good morning,"),
    ("",""),
    ("","My name is Eesh Saxena. I wanted to reach out and enquire if you have any paid internship opportunities available."),
    ("",""),
    ("","I've completed research internships at prestigious institutions, built production web apps with the MERN stack, and worked on ML projects ranging from NLP and computer vision to edge AI deployment. I'm comfortable moving between frontend, backend, and model training depending on what's needed."),
    ("",""),
    ("","I've attached my resume for your reference. Would love to hear from you if there's a fit."),
    ("",""),
    ("","Best,"),
    ("","Eesh Saxena"),
    ("","7976212108 | eeshsaxena@gmail.com"),
]
for i,(lbl,val) in enumerate(template,2):
    ws4.cell(row=i,column=1).value=lbl
    ws4.cell(row=i,column=2).value=val
    ws4.cell(row=i,column=1).font=Font(name="Arial",bold=True,size=9,color="1F3864")
    ws4.cell(row=i,column=2).font=Font(name="Arial",size=10)
    ws4.cell(row=i,column=2).alignment=Alignment(wrap_text=True,vertical="top")
    ws4.row_dimensions[i].height=35 if val and len(val)>80 else 15

try:
    wb.save(r"C:\Users\eeshs\Downloads\Dubai_Internship_Outreach.xlsx")
except PermissionError:
    wb.save(r"C:\Users\eeshs\Downloads\Dubai_Internship_Outreach_v2.xlsx")
    print("Saved as v2 (original was open)")
print("DONE")
