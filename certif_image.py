import requests
import cv2
from PIL import ImageFont, ImageDraw, Image
import numpy as np
import time
import os
from pathlib import Path
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import gspread

################################################################################################################################################################
save_to_gdrive = True #set to True if you want it to upload to google drive as well
automated_on_server = True #set to True only if it runs on some machine every day
process_shopify_orders = True
process_spreadsheet_orders = True
single_order = [False, 4165422612634]
text_colors = [(37, 34, 87, 0), (66, 116, 143,0)] #(42,38,105,0)
script_dir = os.path.dirname(os.path.realpath(__file__))
font_paths = [script_dir+"/fonts/CloisterBlack.ttf", script_dir+"/fonts/DeutschGothicCustom-Regular.ttf", script_dir+"/fonts/LEMONMILK-Bold.otf"]
#################################################################################################################################################################


class Certificate:
    certif_content_def = [
        # LORD1
        {
            "paragraphs": [
                [
                    [["Proclamazione",
                    "Proclamation"], 0, False, True, False, 1, 1.0] # text, font, bold, tighten_width_to_fit, scale_space, color, space_coeff
                ],
                [
                    [["Si conferisce,",
                    "Whereas,"], 0, False, True, False, 1, 1.0]
                ],
                [
                    [["LORD $JR$$NAME$",
                    "LORD $JR$$NAME$"], 1, False, True, False, 0, 2.0]
                ],
                [
                    [["(",
                    "("], 2, False, False, False, 0, 1.0], 
                    [["Il qui di seguito nominato ",
                    "hereafter referred to as "], 0, False, False, False, 0, 1.0], 
                    [["LORD",
                    "THE LORD"], 1, False, False, False, 0, 1.0], 
                    [[")",
                    ")"], 2, False, False, False, 0, 1.0], 
                    [[", avendo, a titolo di avviso, dichiarato l'intenzione di acquistare, e noi Titoli Nobiliari con l'intenzione di accettare, abbiamo acconsentito con il Lord, questo concordato, per la dedica di un lotto di terreno souvenir di $MQ$ in particolare l'appezzamento di terreno identificato e descritto come:",
                    ", having, by way of petition, delivered unto us the intention to purchase, and us with the intention to accept, Titoli Nobiliari has agreed with the Lord, this day, for the sale of a plot of souvenir land of $MQ$ in particular the plot of land identified and described with:"], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["Lotto Numero: §",
                    "The Plot Number: §"], 0, False, False, False, 0, 1.0],
                    [["$NUMBER$",
                    "$NUMBER$"], 0, True, False, False, 0, 1.0]
                ],
                [
                    [["Data: $DATE$",
                    "Date: $DATE$"], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["Per conto di Titoli Nobiliari",
                    "On behalf of Titoli Nobiliari"], 0, False, False, False, 0, 1.0]
                ]
            ],
            "geometries": [
                [(119,), (356, 264), 1765-356, 30, (0,), (1, 140)],
                [(80,), (356, 520), 1808-356, 30, (0,), (1, 94)],
                [(76,), (356, 635), 1808-356, 30, (0,), (1, 89)],
                [(38, 38, 38, 38, 38), (356, 744), 1808-356, 54, (0, 3, -3, 0, 0), (0, 44)],
                [(60, 60), (356, 990), 1338-356, 30, (0,0), (0, 70)],
                [(48,), (356, 1074), 1808-356, 30, (0,), (0, 56)],
                [(37,), (356, 1229), 804-356, 30, (0,), (0, 43)],
            ],
            "template_path": [script_dir+"/certificate_templates/Lord, Lady and Couple #1.png", script_dir+"/certificate_templates/Lord, Lady and Couple #1.png"]
        },

        # LORD2
        {
            "paragraphs": [
                [
                    [["Proclamazione",
                    "Proclamation"], 0, False, True, False, 1, 1.0]
                ],
                [
                    [["Si conferisce,",
                    "Whereas,"], 0, False, True, False, 1, 1.0]
                ],
                [
                    [["LORD $JR$$NAME$",
                    "LORD $JR$$NAME$"], 1, False, True, False, 0, 2.0]
                ],
                [
                    [["(",
                    "("], 2, False, False, False, 0, 1.0], 
                    [["Il qui di seguito nominato ",
                    "hereafter referred to as "], 0, False, False, False, 0, 1.5], 
                    [["LORD",
                    "THE LORD"], 1, False, False, False, 0, 1.0], 
                    [[")",
                    ")"], 2, False, False, False, 0, 1.0], 
                    [[", avendo, a titolo di avviso, questo $MONTHDAY$ giorno di $MONTH$ dell’anno $YEAR$, $RYEAR$ regno della Nostra sovrana Regina Elisabetta Seconda, per grazia Divina, del Regno Unito di Gran Bretagna e Irlanda del Nord, e dei suoi altri Regni e Territori, Regina Sovrana del Commonwealth, Difensore della fede, consegnataci l’intenzione di acquistare, e noi Titoli Nobiliari con l’intenzione di accettare, ha concordato la richiesta del Lord di lasciar loro dedicare una richiesta di possedimento di un lotto di terreno souvenir di precisamente $MQ$, nelle terre incontaminate di Scozia. In particolare il lotto di terra identificato e descritto con la specifica identificativa di lotto numero: §$NUMBER$, di Titoli Nobiliari, della misura di $AREA$ e quindi riferito come “IL LOTTO”. Titoli Nobiliari acconsente di dedicare il lotto, nel nome della Regina nella locazione descritta in questo atto.",
                    ", has, by way of notice, this $MONTHDAY$ day of $MONTH$ in the year $YEAR$, in the $RYEAR$Reign of Our Sovereign Lady Elizabeth the Second, by the Grace of God, of the United Kingdom of Great Britain and North Ireland and of Her other Realms and territories Queen, Head of the Commonwealth, Defender of the Faith, delivered unto us the intention to purchase, and us with the intention to accept, Titoli Nobiliari has agreed with the Lord to bequeath unto them, a dedication of a plot of souvenir land of precisely $MQ$ in West Scotland, in particular, the plot of land identified and described with the specific identifying plot number: §$NUMBER$, by Titoli Nobiliari, of a measurement of $AREA$, and hereafter referred to as ”THE PLOT”. Titoli Nobiliari agrees to dedicate the plot, in the name of the Queen, to the location described in this deed."], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["Riconoscendo il terreno che IL LOTTO forma, contenuto all’interno dell’originario schema dal quale è stato inizialmente identificato e definito come lotto souvenir, Titoli Nobiliari, RICONOSCE tutti gli accordi onorati dalla Regina, dei quali ha preso atto, e riconosciuto un tributo in favore della Regina, i suoi funzionari e successori, tutti loro all’unanime, ma senza i diritti conferiti dai suoi funzionari e successori. La Regina con la presente si impegna con Titoli Nobiliari che il tributo concordato in questa proclamazione è per la Regina e i suoi successori e solo per il loro titolo, e che essi, e nessuno dei loro successori, non venderanno la dedica del lotto numero §$NUMBER$, più specificamente non in modo tale da poter essere registrato o posseduto in titoli separati o in proprietà separate.",
                    "Whereas, part of the estate that THE PLOT forms a part of has been identified and set aside in relation to a scheme of souvenir plots, Titoli Nobiliari, in CONSIDERATION to all sums due and paid to us by the Lady, of which we acknowledge receipt of, has been bequeathed a dedication in favour of the Lady, her assigneers and her successors all and whole, but without the rights thereto over the larger subjects and its successors in title of the larger subjects and all other authorized by it, which remain with the registrants of the estate and larger area and may be identified altogether with the larger area. The Lady hereby covenants with Titoli Nobiliari that the dedication agreed upon in this Proclamation is for the Lady and their successors in title only and that they and any of their successors shall not sell the dedication of plot number §$NUMBER$, more specifically not in such a way that it could be registered or owned in separate titles or in separate ownerships."], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["$NAME$",
                    "$NAME$"], 1, False, False, False, 0, 2.0],
                    [[", in virtù della proprietà di un terreno souvenir in Scozia, a titolo di dedica, e alla ricezione del presente atto, in particolare a riguardo del terreno descritto come lotto numero §$NUMBER$ da Titoli Nobiliari, potrà d'ora in poi e per sempre essere nominato come LORD, e d'ora innanzi, essere conosciuto come ",
                    ", by the virtue of the ownership of land, by way of a dedication, upon the effect and the receipt of this proclamation, in particular regarding the land described as Plot §$NUMBER$ by Titoli Nobiliari, may henceforth and in perpetuity be known by the style and title of LORD, and shall hereafter, to all and sundry, be known as "], 0, False, False, False, 0, 1.0],
                    [["LORD $JR$$NAME$",
                    "LORD $JR$$NAME$"], 1, False, False, False, 0, 2.0]
                ],
                [
                    [["QUESTO ATTO TESTIMONIA IL SEGUENTE",
                    "NOW THIS DEED WITNESSETH AS FOLLOWS"], 1, False, True, False, 1, 2.0]
                ],
                [
                    [["ATTESTATO CHE,",
                    "KNOW YE THEREFORE THAT,"], 1, False, True, False, 1, 2.0]
                ],
                [
                    [["Per conto di Titoli Nobiliari",
                    "On behalf of Titoli Nobiliari"], 0, False, True, False, 0, 1.0]
                ]
            ],
            "geometries": [
                [(115,), (220, 129), 2512-220, 62, (0,), (1, 140)],
                [(52,), (220, 248), 1305-220, 62, (0,), (1, 94)],
                [(52,), (220, 314), 1305-220, 62, (0,), (1, 74)],
                [(36, 36, 36, 36, 36), (220, 378), 1305-220, 43.5, (0, 3, -3, 0, 0), (0, 52)],
                [(36,), (1440, 378), 2525-1440, 43.5, (0,), (0, 52)],
                [(31, 36, 31), (220, 1076), 1675-220, 43.5, (-6,6,0), (0, 52)],
                [(49,), (1440, 318), 2525-1440, 62, (0,), (1, 78)],
                [(40,), (220, 1012), 1675-220, 62, (0,), (1, 94)],
                [(36,), (734, 1286), 1186-734, 62, (0,), (1, 94)],
            ],
            "template_path": [script_dir+"/certificate_templates/Lord, Lady and Couple #2.png", script_dir+"/certificate_templates/ENG-Lord, Lady and Couple #2.png"]
        },


        # LADY1
        {
            "paragraphs": [
                [
                    [["Proclamazione",
                    "Proclamation"], 0, False, True, False, 1, 1.0] # text, font, bold, tighten_width_to_fit, scale_space, color, space_coeff
                ],
                [
                    [["Si conferisce,",
                    "Whereas,"], 0, False, True, False, 1, 1.0]
                ],
                [
                    [["LADY $JR$$NAME$",
                    "LADY $JR$$NAME$"], 1, False, True, False, 0, 2.0]
                ],
                [
                    [["(",
                    "("], 2, False, False, False, 0, 1.0], 
                    [["La qui di seguito nominata ",
                    "hereafter referred to as "], 0, False, False, False, 0, 1.0], 
                    [["LADY",
                    "THE LADY"], 1, False, False, False, 0, 1.0],
                    [[")",
                    ")"], 2, False, False, False, 0, 1.0], 
                    [[", avendo, a titolo di avviso, dichiarato l'intenzione di acquistare, e noi Titoli Nobiliari con l'intenzione di accettare, abbiamo acconsentito con la Lady, questo concordato, per la dedica di un lotto di terreno souvenir di $MQ$ in particolare l'appezzamento di terreno identificato e descritto come:",
                    ", having, by way of petition, delivered unto us the intention to purchase, and us with the intention to accept, Titoli Nobiliari has agreed with the Lady, this day, for the sale of a plot of souvenir land of $MQ$ in particular the plot of land identified and described with:"], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["Lotto Numero: §",
                    "The Plot Number: §"], 0, False, False, False, 0, 1.0],
                    [["$NUMBER$",
                    "$NUMBER$"], 0, True, False, False, 0, 1.0]
                ],
                [
                    [["Data: $DATE$",
                    "Date: $DATE$"], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["Per conto di Titoli Nobiliari",
                    "On behalf of Titoli Nobiliari"], 0, False, False, False, 0, 1.0]
                ]
            ],
            "geometries": [
                [(119,), (356, 264), 1765-356, 30, (0,), (1, 140)],
                [(80,), (356, 520), 1808-356, 30, (0,), (1, 94)],
                [(76,), (356, 635), 1808-356, 30, (0,), (1, 89)],
                [(38, 38, 38, 38, 38), (356, 744), 1808-356, 54, (0, 3, -3, 0, 0), (0, 44)],
                [(60, 60), (356, 990), 1338-356, 30, (0,0), (0, 70)],
                [(48,), (356, 1074), 1808-356, 30, (0,), (0, 56)],
                [(37,), (356, 1229), 804-356, 30, (0,), (0, 43)],
            ],
            "template_path": [script_dir+"/certificate_templates/Lord, Lady and Couple #1.png", script_dir+"/certificate_templates/Lord, Lady and Couple #1.png"]
        },

        # LADY2
        {
            "paragraphs": [
                [
                    [["Proclamazione",
                    "Proclamation"], 0, False, True, False, 1, 1.0]
                ],
                [
                    [["Si conferisce,",
                    "Whereas,"], 0, False, True, False, 1, 1.0]
                ],
                [
                    [["LADY $JR$$NAME$",
                    "LADY $JR$$NAME$"], 1, False, True, False, 0, 2.0]
                ],
                [
                    [["(",
                    "("], 2, False, False, False, 0, 1.0], 
                    [["La qui di seguito nominata ",
                    "hereafter referred to as "], 0, False, False, False, 0, 1.5], 
                    [["LADY",
                    "THE LADY"], 1, False, False, False, 0, 1.0], 
                    [[")",
                    ")"], 2, False, False, False, 0, 1.0], 
                    [[", avendo, a titolo di avviso, questo $MONTHDAY$ giorno di $MONTH$ dell’anno $YEAR$, $RYEAR$ regno della Nostra sovrana Regina Elisabetta Seconda, per grazia Divina, del Regno Unito di Gran Bretagna e Irlanda del Nord, e dei suoi altri Regni e Territori, Regina Sovrana del Commonwealth, Difensore della fede, consegnataci l’intenzione di acquistare, e noi Titoli Nobiliari con l’intenzione di accettare, ha concordato la richiesta della Lady di lasciar loro dedicare una richiesta di possedimento di un lotto di terreno souvenir di precisamente $MQ$, nelle terre incontaminate di Scozia. In particolare il lotto di terra identificato e descritto con la specifica identificativa di lotto numero: §$NUMBER$, di Titoli Nobiliari, della misura di $AREA$ e quindi riferito come “IL LOTTO”. Titoli Nobiliari acconsente di dedicare il lotto, nel nome della Regina nella locazione descritta in questo atto.",
                    ", has, by way of notice, this $MONTHDAY$ day of $MONTH$ in the year $YEAR$, in the $RYEAR$Reign of Our Sovereign Lady Elizabeth the Second, by the Grace of God, of the United Kingdom of Great Britain and North Ireland and of Her other Realms and territories Queen, Head of the Commonwealth, Defender of the Faith, delivered unto us the intention to purchase, and us with the intention to accept, Titoli Nobiliari has agreed with the Lady to bequeath unto them, a dedication of a plot of souvenir land of precisely $MQ$ in West Scotland, in particular, the plot of land identified and described with the specific identifying plot number: §$NUMBER$, by Titoli Nobiliari, of a measurement of $AREA$, and hereafter referred to as ”THE PLOT”. Titoli Nobiliari agrees to dedicate the plot, in the name of the Queen, to the location described in this deed."], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["Riconoscendo il terreno che IL LOTTO forma, contenuto all’interno dell’originario schema dal quale è stato inizialmente identificato e definito come lotto souvenir, Titoli Nobiliari, RICONOSCE tutti gli accordi onorati dalla Regina, dei quali ha preso atto, e riconosciuto un tributo in favore della Regina, i suoi funzionari e successori, tutti loro all’unanime, ma senza i diritti conferiti dai suoi funzionari e successori. La Regina con la presente si impegna con Titoli Nobiliari che il tributo concordato in questa proclamazione è per la Regina e i suoi successori e solo per il loro titolo, e che essi, e nessuno dei loro successori, non venderanno la dedica del lotto numero §$NUMBER$, più specificamente non in modo tale da poter essere registrato o posseduto in titoli separati o in proprietà separate.",
                    "Whereas, part of the estate that THE PLOT forms a part of has been identified and set aside in relation to a scheme of souvenir plots, Titoli Nobiliari, in CONSIDERATION to all sums due and paid to us by the Lady, of which we acknowledge receipt of, has been bequeathed a dedication in favour of the Lady, her assigneers and her successors all and whole, but without the rights thereto over the larger subjects and its successors in title of the larger subjects and all other authorized by it, which remain with the registrants of the estate and larger area and may be identified altogether with the larger area. The Lady hereby covenants with Titoli Nobiliari that the dedication agreed upon in this Proclamation is for the Lady and their successors in title only and that they and any of their successors shall not sell the dedication of plot number §$NUMBER$, more specifically not in such a way that it could be registered or owned in separate titles or in separate ownerships."], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["$NAME$",
                    "$NAME$"], 1, False, False, False, 0, 2.0],
                    [[", in virtù della proprietà di un terreno souvenir in Scozia, a titolo di dedica, e alla ricezione del presente atto, in particolare a riguardo del terreno descritto come lotto numero §$NUMBER$ da Titoli Nobiliari, potrà d'ora in poi e per sempre essere nominata come LADY, e d'ora innanzi, essere conosciuta come ",
                    ", by the virtue of the ownership of land, by way of a dedication, upon the effect and the receipt of this proclamation, in particular regarding the land described as Plot §$NUMBER$ by Titoli Nobiliari, may henceforth and in perpetuity be known by the style and title of LADY, and shall hereafter, to all and sundry, be known as "], 0, False, False, False, 0, 1.0],
                    [["LADY $JR$$NAME$",
                    "LADY $JR$$NAME$"], 1, False, False, False, 0, 2.0]
                ],
                [
                    [["QUESTO ATTO TESTIMONIA IL SEGUENTE",
                    "NOW THIS DEED WITNESSETH AS FOLLOWS"], 1, False, True, False, 1, 2.0]
                ],
                [
                    [["ATTESTATO CHE,",
                    "KNOW YE THEREFORE THAT,"], 1, False, True, False, 1, 2.0]
                ],
                [
                    [["Per conto di Titoli Nobiliari",
                    "On behalf of Titoli Nobiliari"], 0, False, True, False, 0, 1.0]
                ]
            ],
            "geometries": [
                [(115,), (220, 129), 2512-220, 62, (0,), (1, 140)],
                [(52,), (220, 248), 1305-220, 62, (0,), (1, 94)],
                [(52,), (220, 314), 1305-220, 62, (0,), (1, 74)],
                [(36, 36, 36, 36, 36), (220, 378), 1305-220, 43.5, (0, 3, -3, 0, 0), (0, 52)],
                [(36,), (1440, 378), 2525-1440, 43.5, (0,), (0, 52)],
                [(31, 36, 31), (220, 1076), 1675-220, 43.5, (-6,6,0), (0, 52)],
                [(49,), (1440, 318), 2525-1440, 62, (0,), (1, 78)],
                [(40,), (220, 1012), 1675-220, 62, (0,), (1, 94)],
                [(36,), (734, 1286), 1186-734, 62, (0,), (1, 94)],
            ],
            "template_path": [script_dir+"/certificate_templates/Lord, Lady and Couple #2.png", script_dir+"/certificate_templates/ENG-Lord, Lady and Couple #2.png"]
        },


        # COUPLE1
        {
            "paragraphs": [
                [
                    [["Proclamazione",
                    "Proclamation"], 0, False, True, False, 1, 1.0] # text, font, bold, tighten_width_to_fit, scale_space, color, space_coeff
                ],
                [
                    [["Si conferiscono,",
                    "Whereas,"], 0, False, True, False, 1, 1.0]
                ],
                [
                    [["$TITLE1$ $NAME1$",
                    "$TITLE1$ $NAME1$"], 1, False, True, False, 0, 2.0]
                ],
                [
                    [["$TITLE2$ $NAME2$",
                    "$TITLE2$ $NAME2$"], 1, False, True, False, 0, 2.0]
                ],
                [
                    [["(",
                    "("], 2, False, False, False, 0, 1.0], 
                    [["I qui di seguito nominati ",
                    "hereafter referred to as "], 0, False, False, False, 0, 1.5], 
                    [["$TITLE1$ e $TITLE2$",
                    "$TITLE1$ and $TITLE2$"], 1, False, False, False, 0, 1.0], 
                    [[")",
                    ")"], 2, False, False, False, 0, 1.0], 
                    [[", avendo, a titolo di avviso, dichiarato l'intenzione di acquistare, e noi Titoli Nobiliari con l'intenzione di accettare, abbiamo acconsentito con $SPELLA$, questo concordato, per la dedica di un lotto di terreno souvenir di $MQ$ in particolare l'appezzamento di terreno identificato e descritto come:",
                    ", having, by way of petition, delivered unto us the intention to purchase, and us with the intention to accept, Titoli Nobiliari has agreed with $SPELLA$, this day, for the sale of a plot of souvenir land of $MQ$ in particular the plot of land identified and described with:"], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["Lotto Numero: §",
                    "The Plot Number: §"], 0, False, False, False, 0, 1.0],
                    [["$NUMBER$",
                    "$NUMBER$"], 0, True, False, False, 0, 1.0]
                ],
                [
                    [["Data: $DATE$",
                    "Date: $DATE$"], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["Per conto di Titoli Nobiliari",
                    "On behalf of Titoli Nobiliari"], 0, False, False, False, 0, 1.0]
                ]
            ],
            "geometries": [
                [(119,), (356, 220), 1765-356, 30, (0,), (1, 140)],
                [(80,), (356, 405), 1808-356, 30, (0,), (1, 94)],
                [(76,), (356, 520), 1808-356, 30, (0,), (1, 89)],
                [(76,), (356, 635), 1808-356, 30, (0,), (1, 89)],
                [(38, 38, 38, 38, 38), (356, 744), 1808-356, 54, (0, 3, -3, 0, 0), (0, 44)],
                [(60, 60), (356, 990), 1338-356, 30, (0,0), (0, 70)],
                [(48,), (356, 1074), 1808-356, 30, (0,), (0, 56)],
                [(37,), (356, 1229), 804-356, 30, (0,), (0, 43)],
            ],
            "template_path": [script_dir+"/certificate_templates/Lord, Lady and Couple #1.png", script_dir+"/certificate_templates/Lord, Lady and Couple #1.png"]
        },

        # COUPLE2
        {
            "paragraphs": [
                [
                    [["Proclamazione",
                    "Proclamation"], 0, False, True, False, 1, 1.0]
                ],
                [
                    [["Si conferiscono,",
                    "Whereas,"], 0, False, True, False, 1, 1.0]
                ],
                [
                    [["$TITLE1$ $NAME1$ & $TITLE2$ $NAME2$",
                    "$TITLE1$ $NAME1$ & $TITLE2$ $NAME2$"], 1, False, True, False, 0, 2.0]
                ],
                [
                    [["(",
                    "("], 2, False, False, False, 0, 1.0], 
                    [["I qui di seguito nominati ",
                    "hereafter referred to as "], 0, False, False, False, 0, 2.0], 
                    [["$TITLE1$ e $TITLE2$",
                    "$TITLE1$ and $TITLE2$"], 1, False, False, False, 0, 1.0], 
                    [[")",
                    ")"], 2, False, False, False, 0, 1.0],
                    [[", avendo, a titolo di avviso, questo $MONTHDAY$ giorno di $MONTH$ dell’anno $YEAR$, $RYEAR$ regno della Nostra sovrana Regina Elisabetta Seconda, per grazia Divina, del Regno Unito di Gran Bretagna e Irlanda del Nord, e dei suoi altri Regni e Territori, Regina Sovrana del Commonwealth, Difensore della fede, consegnataci l’intenzione di acquistare, e noi Titoli Nobiliari con l’intenzione di accettare, ha concordato la richiesta $SPELLB$ di lasciar loro dedicare una richiesta di possedimento di un lotto di terreno di precisamente $MQ$, nelle terre incontaminate di Scozia. In particolare il lotto di terra identificato e descritto con la specifica identificativa di lotto numero: §$NUMBER$, di Titoli Nobiliari, della misura di $AREA$ e quindi riferito come “IL LOTTO”. Titoli Nobiliari acconsente di dedicare il lotto, nel nome della Regina nella locazione descritta in questo atto.",
                    ", have, by way of notice, this $MONTHDAY$ day of $MONTH$ in the year $YEAR$, in the $RYEAR$Reign of Our Sovereign Lady Elizabeth the Second, by the Grace of God, of the United Kingdom of Great Britain and North Ireland and of Her other Realms and territories Queen, Head of the Commonwealth, Defender of the Faith, delivered unto us the intention to purchase, and us with the intention to accept, Titoli Nobiliari has agreed with $SPELLB$ to bequeath unto them, a dedication of a plot of souvenir land of precisely $MQ$ in West Scotland, in particular, the plot of land identified and described with the specific identifying plot number: §$NUMBER$, by Titoli Nobiliari, of a measurement of $AREA$, and hereafter referred to as ”THE PLOT”. Titoli Nobiliari agrees to dedicate the lot, in the name of the Queen, to the location described in this deed."], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["Riconoscendo il terreno che IL LOTTO forma, contenuto all’interno dell’originario schema dal quale è stato inizialmente identificato e definito come lotto souvenir, Titoli Nobiliari, RICONOSCE tutti gli accordi onorati dalla Regina, dei quali ha preso atto, e riconosciuto un tributo in favore della Regina, i suoi funzionari e successori, tutti loro all’unanime, ma senza i diritti conferiti dai suoi funzionari e successori. La Regina con la presente si impegna con Titoli Nobiliari che il tributo concordato in questa proclamazione è per la Regina e i suoi successori e solo per il loro titolo, e che essi, e nessuno dei loro successori, non venderanno la dedica del lotto numero §$NUMBER$, più specificamente non in modo tale da poter essere registrato o posseduto in titoli separati o in proprietà separate.",
                    "Whereas, part of the estate that THE PLOT forms a part of has been identified and set aside in relation to a scheme o souvenir plots, Titoli Nobiliari, in CONSIDERATION to all sums due and paid to us by the Lady, of which we acknowledge receipt of, has been bequeathed a dedication in favour of the Lady, her assigneers and her successors all and whole, but without the rights thereto over the larger subjects and its successors in title of the larger subjects and all other authorized by it, which remain with the registrants of the estate and larger area and may be identified altogether with the larger area. The Lady hereby covenants with Titoli Nobiliari that the dedication agreed upon in this Proclamation is for the Lady and their successors in title only and that they and any of their successors shall not sell the dedication of plot number §$NUMBER$, more specifically not in such a way that it could be registered or owned in separate titles or in separate ownerships."], 0, False, False, False, 0, 1.0]
                ],
                [
                    [["$NAME1$ e $NAME2$",
                    "$NAME1$ and $NAME2$"], 1, False, False, False, 0, 2.0],
                    [[", in virtù della proprietà di un terreno souvenir in Scozia, a titolo di dedica, e alla ricezione del presente atto, in particolare a riguardo del terreno descritto come lotto numero §$NUMBER$ da Titoli Nobiliari, potranno d'ora in poi e per sempre essere nominati come $TITLE1$ e $TITLE2$, e d'ora innanzi, essere conosciuti come ",
                    ", by the virtue of the ownership of land, by way of a dedication, upon the effect and the receipt of this proclamation, in particular regarding the land described as Plot §$NUMBER$ by Titoli Nobiliari, may henceforth and in perpetuity be known by the style and title of $TITLE1$ and $TITLE2$, and shall hereafter, to all and sundry, be known as "], 0, False, False, False, 0, 1.0],
                    [["$TITLE1$ $NAME1$ & $TITLE2$ $NAME2$",
                    "$TITLE1$ $NAME1$ & $TITLE2$ $NAME2$"], 1, False, False, True, 0, 2.0]
                ],
                [
                    [["QUESTO ATTO TESTIMONIA IL SEGUENTE",
                    "NOW THIS DEED WITNESSETH AS FOLLOWS"], 1, False, True, False, 1, 2.0]
                ],
                [
                    [["ATTESTATO CHE,",
                    "KNOW YE THEREFORE THAT,"], 1, False, True, False, 1, 2.0]
                ],
                [
                    [["Per conto di Titoli Nobiliari",
                    "On behalf of Titoli Nobiliari"], 0, False, True, False, 0, 1.0]
                ]
            ],
            "geometries": [
                [(115,), (220, 90), 2512-220, 62, (0,), (1, 140)],
                [(52,), (220, 175), 1305-220, 62, (0,), (1, 94)],
                [(51,), (220, 250), 2512-220, 62, (0,), (1, 74)],
                [(36, 36, 36, 36, 36), (220, 325), 1305-220, 43.5, (0, 3, -3, 0, 0), (0, 52)],
                [(36,), (1440, 378), 2525-1440, 43.5, (0,), (0, 52)],
                [(31, 36, 31), (220, 1026), 1820-220, 43.5, (-6,6,0), (0, 52)],
                [(40,), (1440, 330), 2525-1440, 62, (0,), (1, 78)],
                [(40,), (220, 962), 1675-220, 62, (0,), (1, 94)],
                [(36,), (734, 1286), 1186-734, 62, (0,), (1, 94)],
            ],
            "template_path": [script_dir+"/certificate_templates/Lord, Lady and Couple #2.png", script_dir+"/certificate_templates/ENG-Lord, Lady and Couple #2.png"]
        },
    ]

    def __init__(self, certif_type, fill_params, item_lang) -> None: #type = {0,1,2,3,4,5}
        self.certif_type = certif_type
        self.fill_params = fill_params
        self.item_lang = item_lang

        self.img_pil = Image.fromarray(cv2.imread(self.certif_content_def[certif_type]["template_path"][self.item_lang]))
        self.draw = ImageDraw.Draw(self.img_pil)
        

    # Font can even be NON MONOSPACE
    def justified_text_to_image(self, offset, pos, multiline_text, font, line_width, line_space, bold, scale_space, center, color_text, base_space_koef):
        words = multiline_text.split(' ')
        space_char_width = font.getsize(" ")[0] * base_space_koef
        words_width = { words[i]: font.getsize(words[i])[0] for i in range(len(words)) }
        
        line_words = [[]]
        line_words_width_sum = []
        line = 0
        sum_width = 0
        i = 0
        while i < len(words):
            if sum_width + words_width[words[i]] <= (line_width if line!=0 else line_width-offset[0]):
                sum_width += words_width[words[i]] + space_char_width
                line_words[line].append(words[i])
                i += 1
            if i==len(words) or sum_width + words_width[words[i]] > (line_width if line!=0 else line_width-offset[0]):
                sum_width -= space_char_width #last space del
                line_words_width_sum.append(sum_width)
                if i<len(words): line_words.append([])
                line += 1
                sum_width = 0
        
        last_offset = offset
        for line in range(len(line_words)):
            line_spaces = len(line_words[line])-1
            if scale_space and line_spaces>0:
                space_multiplier = ((line_width if line!=0 else line_width-offset[0]) - (line_words_width_sum[line] - line_spaces*space_char_width)) / (line_spaces*space_char_width) if line < len(line_words)-1 else 1
            else:
                space_multiplier = 1
            x_last = pos[0] + (offset[0] if line==0 else 0) + ((line_width - line_words_width_sum[0])/2 if center else 0)
            y = pos[1] + line*line_space + offset[1]
            for word in line_words[line]:
                self.draw.text((x_last, y), word, font=font, fill=color_text, stroke_width=(1 if bold else 0))
                x_last += words_width[word] + space_multiplier*space_char_width
            last_offset = [line_words_width_sum[line], line*line_space]
        return last_offset, len(line_words)<=1

    
    def fill_placeholders(self, text_piece):
        for fill in self.fill_params:
            text_piece = text_piece.replace(fill[0], fill[1])
        return text_piece


    def common_couple_name_fontsize(self, text_piece1, text_piece2, defaut_fontsize, max_width):
        text_piece1_filled = self.fill_placeholders(text_piece1[0][self.item_lang])
        text_piece2_filled = self.fill_placeholders(text_piece2[0][self.item_lang])
        
        font = ImageFont.truetype(font_paths[text_piece1[1]], defaut_fontsize)
        fnt_decr = 0
        spc1 = text_piece1_filled.split(" "); spc2 = text_piece2_filled.split(" ")
        letters1 = ''.join(spc1); letters2 = ''.join(spc2)

        while text_piece1[3] and text_piece2[3] and ( (font.getsize(letters1)[0] + font.getsize(" ")[0]*len(spc1)*text_piece1[6]) > 0.95*max_width or (font.getsize(letters2)[0] + font.getsize(" ")[0]*len(spc2)*text_piece2[6]) > 0.95*max_width):
            fnt_decr += 1
            font = ImageFont.truetype(font_paths[text_piece1[1]], defaut_fontsize - fnt_decr)
        
        return defaut_fontsize - fnt_decr


    def make_certificate_image(self) -> Image:
        content_def = self.certif_content_def[self.certif_type]
        paragraphs = content_def["paragraphs"]
        geometries = content_def["geometries"]

        if self.certif_type == 4: #couple certif 1
            cople_name_fontsize = self.common_couple_name_fontsize(paragraphs[2][0], paragraphs[3][0], geometries[2][0][0], geometries[2][2])

        for i in range(len(paragraphs)):
            offset = [0, 0 + geometries[i][4][0]]
            for j in range(len(paragraphs[i])):
                text_piece = paragraphs[i][j]
                text_piece_filled = self.fill_placeholders(text_piece[0][self.item_lang])
                
                if self.certif_type == 4 and (i in [2, 3]):
                    font = ImageFont.truetype(font_paths[text_piece[1]], cople_name_fontsize)
                else:
                    font = ImageFont.truetype(font_paths[text_piece[1]], geometries[i][0][j])
                    fnt_decr = 0
                    spc = text_piece_filled.split(" ")
                    letters = ''.join(spc)
                    while text_piece[3] and (font.getsize(letters)[0] + font.getsize(" ")[0]*len(spc)*text_piece[6]) > 0.95*geometries[i][2]:
                        fnt_decr += 1
                        font = ImageFont.truetype(font_paths[text_piece[1]], geometries[i][0][j] - fnt_decr)

                new_offset, one_line = self.justified_text_to_image(offset, geometries[i][1], text_piece_filled, font, geometries[i][2], geometries[i][3], text_piece[2], not(text_piece[4]), (i==0 and self.certif_type in [0,1,2,3,4,5] or i==2 and self.certif_type in [5]), text_colors[text_piece[5]], text_piece[6])
                if not one_line: offset[0] = 0
                offset[0] += new_offset[0]
                offset[1] += new_offset[1] + geometries[i][4][j]
        return self.img_pil


    def save_certificate_image(self, filename, order_number, save_to_gdrive, google_folder_id) -> None:
        img = np.array(self.img_pil)
        folder = "{0}/orders".format(script_dir) if automated_on_server else "{0}/orders/order-{1}".format(script_dir, order_number)
        Path(folder).mkdir(parents=True, exist_ok=True)
        filepath = "{0}/{1}.png".format(folder, filename)
        cv2.imwrite(filepath, img)
        
        if save_to_gdrive:
            gfile = drive.CreateFile({'title': filename, 'parents': [{'id': google_folder_id}]})
            gfile.SetContentFile(filepath)
            gfile.Upload()
            if automated_on_server:
                os.remove(filepath)






shopify_product_ids = {
    6793463562394: "LORD",
    6793677701274: "LADY",
    6805431877786: "COUPLE",
    6902503866522: "FAMILIA"
}
monthdays = [["primo", "first"], ["secondo", "second"], ["terzo", "third"], ["quarto", "fourth"], ["quinto", "fifth"], ["sesto", "sixth"], ["settimo", "seventh"],
            ["ottavo", "eighth"], ["nono", "ninth"], ["decimo", "tenth"], ["undicesimo", "eleventh"], ["dodicesimo", "twelfth"], ["tredicesimo", "thirteenth"], 
            ["quattordicesimo", "fourteenth"], ["quindicesimo", "fifteenth"], ["sedicesimo", "sixteenth"], ["diciassettesimo", "seventeenth"], ["diciottesimo", "eighteenth"],
            ["diciannovesimo", "nineteenth"], ["ventesimo", "twentieth"], ["ventunesimo", "twenty first"], ["ventiduesimo", "twenty second"], ["ventitreesimo", "twenty third"],
            ["ventiquattresimo", "twenty fourth"], ["venticinquesimo", "twenty fifth"], ["ventiseiesimo", "twenty sixth"], ["ventisettesimo", "twenty seventh"], 
            ["ventotto", "twenty eighth"], ["ventinovesimo", "twenty ninth"], ["trentesimo", "thirtieth"], ["trentunesimo", "thirty first"]]
months = [["Gennaio", "January"], ["Febbraio", "February"], ["Marzo", "March"], ["Aprile", "April"], ["Maggio", "May"], ["Giugno", "June"], ["Luglio", "July"], 
            ["Agosto", "August"], ["Settembre", "September"], ["Ottobre", "October"], ["Novembre", "November"], ["Dicembre", "December"]]
mqs = {"1mq": ["un metro quadro", "one square foot"],
        "5mq": ["cinque metri quadri", "five square feet"],
        "10mq": ["dieci metri quadri", "ten square feet"],
        "2mq": ["due metri quadri", "placeholder"],
        "20mq": ["venti metri quadri", "placeholder"]}
areas = {"1mq": ["un metro per un metro", "one foot by one foot"], 
            "5mq": ["cinque metri per un metro", "five feet by one foot"], 
            "10mq": ["cinque metri per due metri", "five feet by two feet"], 
            "2mq": ["due metri per un metro", "placeholder"], 
            "20mq": ["cinque metri per quattro metri", "placeholder"]}
years = {"1953": ["millenovecento e cinquantatre", "one thousand nine hundred and fifty three"], "1954": ["millenovecento e cinquantaquattro", "one thousand nine hundred and fifty four"], 
            "1955": ["millenovecento e cinquantacinque", "one thousand nine hundred and fifty five"], "1956": ["millenovecento e cinquantasei", "one thousand nine hundred and fifty six"], "1957": ["millenovecento e cinquantasette", "one thousand nine hundred and fifty seven"], 
            "1958": ["millenovecento e cinquantotto", "one thousand nine hundred and fifty eight"], "1959": ["millenovecento e cinquantanove", "one thousand nine hundred and fifty nine"], "1960": ["millenovecento e sessanta", "one thousand nine hundred and sixty"], 
            "1961": ["millenovecento e sessantuno", "one thousand nine hundred and sixty one"], "1962": ["millenovecento e sessantadue", "one thousand nine hundred and sixty two"], "1963": ["millenovecento e sessantatre", "one thousand nine hundred and sixty three"], 
            "1964": ["millenovecento e sessantaquattro", "one thousand nine hundred and sixty four"], "1965": ["millenovecento e sessantacinque", "one thousand nine hundred and sixty five"], "1966": ["millenovecento e sessantasei", "one thousand nine hundred and sixty six"], 
            "1967": ["millenovecento e sessantasette", "one thousand nine hundred and sixty seven"], "1968": ["millenovecento e sessantotto", "one thousand nine hundred and sixty eight"], "1969": ["millenovecento e sessantanove", "one thousand nine hundred and sixty nine"], 
            "1970": ["millenovecento e settanta", "one thousand nine hundred and seventy"], "1971": ["millenovecento e settantuno", "one thousand nine hundred and seventy one"], "1972": ["millenovecento e settantadue", "one thousand nine hundred and seventy two"], 
            "1973": ["millenovecento e settantatre", "one thousand nine hundred and seventy three"], "1974": ["millenovecento e settantaquattro", "one thousand nine hundred and seventy four"], "1975": ["millenovecento e settantacinque", "one thousand nine hundred and seventy five"], 
            "1976": ["millenovecento e settantasei", "one thousand nine hundred and seventy six"], "1977": ["millenovecento e settantasette", "one thousand nine hundred and seventy seven"], "1978": ["millenovecento e settantotto", "one thousand nine hundred and seventy eight"], 
            "1979": ["millenovecento e settantanove", "one thousand nine hundred and seventy nine"], "1980": ["millenovecento e ottanta", "one thousand nine hundred and eighty"], "1981": ["millenovecento e ottantuno", "one thousand nine hundred and eighty one"], 
            "1982": ["millenovecento e ottantadue", "one thousand nine hundred and eighty two"], "1983": ["millenovecento e ottantatre", "one thousand nine hundred and eighty three"], "1984": ["millenovecento e ottantaquattro", "one thousand nine hundred and eighty four"], 
            "1985": ["millenovecento e ottantacinque", "one thousand nine hundred and eighty five"], "1986": ["millenovecento e ottantasei", "one thousand nine hundred and eighty six"], "1987": ["millenovecento e ottantasette", "one thousand nine hundred and eighty seven"], 
            "1988": ["millenovecento e ottantotto", "one thousand nine hundred and eighty eight"], "1989": ["millenovecento e ottantanove", "one thousand nine hundred and eighty nine"], "1990": ["millenovecento e novanta", "one thousand nine hundred and ninety"], 
            "1991": ["millenovecento e novantuno", "one thousand nine hundred and ninety one"], "1992": ["millenovecento e novantadue", "one thousand nine hundred and ninety two"], "1993": ["millenovecento e novantatre", "one thousand nine hundred and ninety three"], 
            "1994": ["millenovecento e novantaquattro", "one thousand nine hundred and ninety four"], "1995": ["millenovecento e novantacinque", "one thousand nine hundred and ninety five"], "1996": ["millenovecento e novantasei", "one thousand nine hundred and ninety six"], 
            "1997": ["millenovecento e novantasette", "one thousand nine hundred and ninety seven"], "1998": ["millenovecento e novantotto", "one thousand nine hundred and ninety eight"], "1999": ["millenovecento e novanta", "one thousand nine hundred and ninety nine"], 
            "2000": ["duemila", "two thousand"], "2001": ["duemila e uno", "two thousand and one"], "2002": ["duemila e due", "two thousand and two"], 
            "2003": ["duemila e tre", "two thousand and three"], "2004": ["duemila e quattro", "two thousand and four"], "2005": ["duemila e cinque", "two thousand and five"], 
            "2006": ["duemila e sei", "two thousand and six"], "2007": ["duemila e sette", "two thousand and seven"], "2008": ["duemila e otto", "two thousand and eight"], 
            "2009": ["duemila e nove", "two thousand and nine"], "2010": ["duemila e dieci", "two thousand and ten"], "2011": ["duemila e undici", "two thousand and eleven"], 
            "2012": ["duemila e docici", "two thousand and twelve"], "2013": ["duemila e tredici", "two thousand and thirteen"], "2014": ["duemila e quattordici", "two thousand and fourteen"], 
            "2015": ["duemila e quindici", "two thousand and fifteen"], "2016": ["duemila e sedici", "two thousand and sixteen"], "2017": ["duemila e diciassette", "two thousand and seventeen"], 
            "2018": ["duemila e diciotto", "two thousand and eighteen"], "2019": ["duemila e diciannove", "two thousand and nineteen"], "2020": ["duemila e venti", "two thousand and twenty"], 
            "2021": ["duemila e ventuno", "two thousand and twenty one"], "2022": ["duemila e ventidue", "two thousand and twenty two"], 
            "2023": ["duemila e ventitre", "two thousand and twenty three"], "2024": ["duemila e ventiquattro", "two thousand and twenty four"], 
            "2025": ["duemila e venticinque", "two thousand and twenty five"], "2026": ["duemila e ventisei", "two thousand and twenty six"], 
            "2027": ["duemila e ventisette", "two thousand and twenty seven"], "2028": ["duemila e ventotto", "two thousand and twenty eight"], 
            "2029": ["duemila e ventinove", "two thousand and twenty nine"], "2030": ["duemila e trenta", "two thousand and thirty"], 
            "2031": ["duemila e trentuno", "two thousand and thirty one"]}
ryears = {
            "1953": ["uno", "first"], "1954": ["due", "second"], "1955": ["tre", "third"], "1956": ["quattro", "fourth"], "1957": ["cinque", "fifth"], "1958": ["sei", "sixth"], 
            "1959": ["sette", "seventh"], "1960": ["otto", "eighth"], "1961": ["nove", "ninth"], "1962": ["dieci", "tenth"], "1963": ["undici", "eleventh"], "1964": ["dodici", "twelfth"], 
            "1965": ["tredici", "thirteenth"], "1966": ["quattordici", "fourteenth"], "1967": ["quindici", "fifteenth"], "1968": ["sedici", "sixteenth"], "1969": ["diciassette", "seventeenth"], 
            "1970": ["diciotto", "eighteenth"], "1971": ["diciannove", "nineteenth"], "1972": ["venti", "twentieth"], "1973": ["ventuno", "twenty first"], "1974": ["ventidue", "twenty second"], 
            "1975": ["ventitre", "twenty third"], "1976": ["ventiquattro", "twenty fourth"], "1977": ["venticinque", "twenty fifth"], "1978": ["ventisei", "twenty sixth"], "1979": ["ventisette", "twenty seventh"], 
            "1980": ["ventotto", "twenty eighth"], "1981": ["ventinove", "twenty ninth"], "1982": ["trenta", "thirtieth"], "1983": ["trentuno", "thirty first"], "1984": ["trentadue", "thirty second"], 
            "1985": ["trentatre", "thirty third"], "1986": ["trentaquattro", "thirty fourth"], "1987": ["trentacinque", "thirty fifth"], "1988": ["trentasei", "thirty sixth"], "1989": ["trentasette", "thirty seventh"], 
            "1990": ["trentotto", "thirty eighth"], "1991": ["trentanove", "thirty ninth"], "1992": ["quaranta", "fortieth"], "1993": ["quarantuno", "forty first"], "1994": ["quarantadue", "forty second"], 
            "1995": ["quarantatre", "forty third"], "1996": ["quarantaquattro", "forty fourth"], "1997": ["quarantacinque", "forty fifth"], "1998": ["quarantasei", "forty sixth"], "1999": ["quarantasette", "forty seventh"], 
            "2000": ["quarantotto", "forty eighth"], "2001": ["quarantanove", "forty ninth"], "2002": ["cinquanta", "fiftieth"], "2003": ["cinquantuno", "fifty first"], "2004": ["cinquantadue", "fifty second"], 
            "2005": ["cinquantatre", "fifty third"], "2006": ["cinquantaquattro", "fifty fourth"], "2007": ["cinquantacinque", "fifty fifth"], "2008": ["cinquantasei", "fifty sixth"], "2009": ["cinquantasette", "fifty seventh"], 
            "2010": ["cinquantotto", "fifty eighth"], "2011": ["cinquantanove", "fifty ninth"], "2012": ["sessanta", "sixtieth"], "2013": ["sessantuno", "sixty first"], "2014": ["sessantadue", "sixty second"], 
            "2015": ["sessantatre", "sixty third"], "2016": ["sessantaquattro", "sixty fourth"], "2017": ["sessantacinque", "sixty fifth"], "2018": ["sessantasei", "sixty sixth"], "2019": ["sessantasette", "sixty seventh"], "2020": ["sessantotto", "sixty eighth"],
            "2021": ["sessantanove", "sixty ninth"], "2022": ["settanta", "seventieth"], "2023": ["settantuno", "seventy first"], "2024": ["settantadue", "seventy second"], 
            "2025": ["settantatre", "seventy third"], "2026": ["settantaquattro", "seventy fourth"], "2027": ["settantacinque", "seventy fifth"], 
            "2028": ["settantasei", "seventy sixth"], "2029": ["settantasette", "seventy seventh"], "2030": ["settantotto", "seventy eighth"], "2031": ["settantanove", "seventy ninth"]}


with open(script_dir+"/shopify_apikey.txt", 'r') as f:
    shopify_apikey = f.read().strip()

if save_to_gdrive:
    client_json_path = script_dir+'/client_secrets.json'
    GoogleAuth.DEFAULT_SETTINGS['client_config_file'] = client_json_path
    gauth = GoogleAuth(script_dir+"/settings.yaml")
    drive = GoogleDrive(gauth)

if process_shopify_orders:
    with open(script_dir+"/next_ord_num.txt", 'r') as f:
        number = int(f.read())

    with open(script_dir+"/processed_order_ids.txt", 'r') as f:
        processed_order_ids = list(filter(lambda x: len(x)!=0, f.read().split('\n')))

    if single_order[0]:
        orders = [requests.get(
            "https://{0}@titoli-nobiliari.myshopify.com/admin/api/2021-07/orders/{1}.json?fields=id,order_number,line_items"
            .format(shopify_apikey, single_order[1])).json()["order"]]
    else:
        orders = requests.get(
            "https://{0}@titoli-nobiliari.myshopify.com/admin/api/2021-07/orders.json?since_id={1}&limit={2}&fields=id,order_number,line_items"
            .format(shopify_apikey, processed_order_ids[-1], 250)).json()["orders"]

    #orders.reverse()

    print("Image processing of unfulfilled Shopify orders is starting...")
    k = 0
    for order in orders:
        order_id = order["id"]
        order_number = order["order_number"]
        if not(single_order[0]) and str(order_id) in processed_order_ids:
            continue

        for i in range(len(order["line_items"])):
            item = order["line_items"][i]
            product_id = item["product_id"]
            if not (product_id in shopify_product_ids): continue
            product_type = shopify_product_ids[product_id]
            
            item_lang = 0
            date = time.strftime("%d %m %Y", time.localtime())
            for prop in item["properties"]:
                if prop["name"] == "Lingua" and prop["value"] == "Inglese":
                    item_lang = 1
                if prop["name"] == "Data richiesta":
                    date = ' '.join(prop["value"].split('/'))
            
            
            """
            "nell’anno sessantanove del"
            "nel"

            "sixty ninth year of the "
            ""
            """
            if int(date.split(' ')[2]) > 2031: continue
            try:
                ryearspec = ["nell’anno {0} del".format(ryears[date.split(' ')[2]][item_lang]), "{0} year of the ".format(ryears[date.split(' ')[2]][item_lang])][item_lang]
            except:
                ryearspec = ["nel", ""][item_lang]


            if product_type == "LORD" or product_type == "LADY": #Lord or Lady
                is_lord = product_type == "LORD"

                try:
                    name = ' '.join(item["properties"][0]["value"].upper().strip().split())
                    if len(name) == 0:
                        raise
                    
                    vtitle = item["variant_title"].split(' ')
                    if vtitle[1] == "mq":
                        framesize = vtitle[0]+"mq"
                    else:
                        raise

                    template_fills = [("$NAME$", name), ("$DATE$", date), ("$MQ$", mqs[framesize][item_lang]), ("$NUMBER$", str(number)), ("$MONTHDAY$", monthdays[int(date.split(' ')[0])-1][item_lang]), 
                        ("$MONTH$", months[int(date.split(' ')[1])-1][item_lang]), ("$YEAR$", years[date.split(' ')[2]][item_lang]), ("$RYEAR$", ryearspec), ("$AREA$", areas[framesize][item_lang]),
                        ("$JR$", "")]
                except:
                    continue
                
                cert = Certificate(0 if is_lord else 2, template_fills, item_lang)
                cert.make_certificate_image()
                filename = "#{0}-{1}-A".format(order_number, i+1)
                cert.save_certificate_image(filename, order_number, save_to_gdrive, '1udPWiIwHkgSJKIidTv_EzfFwNAVqeLxf')

                cert = Certificate(1 if is_lord else 3, template_fills, item_lang)
                cert.make_certificate_image()
                filename = "#{0}-{1}-B".format(order_number, i+1)
                cert.save_certificate_image(filename, order_number, save_to_gdrive, '1udPWiIwHkgSJKIidTv_EzfFwNAVqeLxf')
                
                number += 1
                with open(script_dir+"/next_ord_num.txt", 'w') as f:
                    f.write(str(number))
            
            elif product_type == "FAMILIA": # Familia
                try:
                    vtitle = item["variant_title"].split('/')[1].strip().split(' ')
                    if vtitle[1] == "mq":
                        framesize = vtitle[0]+"mq"
                    else:
                        raise
                except:
                    continue

                j = 0
                while item["properties"][j]["name"].find("Titolo") >= 0:
                    try:
                        name = ' '.join(item["properties"][j+1]["value"].upper().strip().split())
                        if len(name) == 0:
                            raise
                        title = item["properties"][j]["value"].strip()
                        template_fills = [("$NAME$", name), ("$DATE$", date), ("$MQ$", mqs[framesize][item_lang]), ("$NUMBER$", str(number)), ("$MONTHDAY$", monthdays[int(date.split(' ')[0])-1][item_lang]), 
                            ("$MONTH$", months[int(date.split(' ')[1])-1][item_lang]), ("$YEAR$", years[date.split(' ')[2]][item_lang]), ("$RYEAR$", ryearspec), ("$AREA$", areas[framesize][item_lang]), 
                            ("$JR$", "JR. " if title.find("Jr.") >= 0 else "")]
                    except:
                        j += 2
                        continue
                    
                    if title == "Lord (padre)" or title == "Lady (madre)" or title.find("Jr.") >= 0: # Parents or Kids
                        is_lord = title.find("Lord") >= 0
                        cert = Certificate(0 if is_lord else 2, template_fills, item_lang)
                        cert.make_certificate_image()
                        filename = "#{0}-{1}-{2}A".format(order_number, i+1, j//2+1)
                        cert.save_certificate_image(filename, order_number, save_to_gdrive, '1udPWiIwHkgSJKIidTv_EzfFwNAVqeLxf')

                        cert = Certificate(1 if is_lord else 3, template_fills, item_lang)
                        cert.make_certificate_image()
                        filename = "#{0}-{1}-{2}B".format(order_number, i+1, j//2+1)
                        cert.save_certificate_image(filename, order_number, save_to_gdrive, '1udPWiIwHkgSJKIidTv_EzfFwNAVqeLxf')
                        
                        number += 1
                        with open(script_dir+"/next_ord_num.txt", 'w') as f:
                            f.write(str(number))
                    j += 2
            
            elif product_type == "COUPLE": # Couple
                try:
                    name1 = ' '.join(item["properties"][2]["value"].upper().strip().split())
                    name2 = ' '.join(item["properties"][3]["value"].upper().strip().split())
                    is_name1_lord = item["properties"][2]["name"].find("Lord") >= 0
                    is_name2_lord = item["properties"][3]["name"].find("Lord") >= 0

                    if len(name1) == 0 or len(name2) == 0:
                        raise
                    
                    vtitle = item["variant_title"].split(' ')
                    if vtitle[1] == "mq":
                        framesize = vtitle[0]+"mq"
                    else:
                        raise
                    
                    if is_name1_lord:
                        if is_name2_lord:
                            spella = ["il Lord e il Lord", "the Lord and the Lord"][item_lang]
                            spellb = ["del Lord e del Lord", "the Lord and the Lord"][item_lang]
                        else:
                            spella = ["il Lord e la Lady", "the Lord and the Lady"][item_lang]
                            spellb = ["del Lord e della Lady", "the Lord and the Lady"][item_lang]
                    else:
                        if is_name2_lord:
                            spella = ["la Lady e il Lord", "the Lady and the Lord"][item_lang]
                            spellb = ["della Lady e del Lord", "the Lady and the Lord"][item_lang]
                        else:
                            spella = ["la Lady e la Lady", "the Lady and the Lady"][item_lang]
                            spellb = ["della Lady e della Lady", "the Lady and the Lady"][item_lang]

                    template_fills = [("$TITLE1$", "LORD" if is_name1_lord else "LADY"), ("$NAME1$", name1), ("$TITLE2$", "LORD" if is_name2_lord else "LADY"), ("$NAME2$", name2), 
                                    ("$DATE$", date), ("$MQ$", mqs[framesize][item_lang]), ("$NUMBER$", str(number)), ("$MONTHDAY$", monthdays[int(date.split(' ')[0])-1][item_lang]), 
                                    ("$MONTH$", months[int(date.split(' ')[1])-1][item_lang]), ("$YEAR$", years[date.split(' ')[2]][item_lang]), ("$RYEAR$", ryearspec), 
                                    ("$AREA$", areas[framesize][item_lang]), ("$SPELLA$", spella), ("$SPELLB$", spellb)]
                except:
                    continue
                
                cert = Certificate(4, template_fills, item_lang)
                cert.make_certificate_image()
                filename = "#{0}-{1}-A".format(order_number, i+1)
                cert.save_certificate_image(filename, order_number, save_to_gdrive, '1udPWiIwHkgSJKIidTv_EzfFwNAVqeLxf')

                cert = Certificate(5, template_fills, item_lang)
                cert.make_certificate_image()
                filename = "#{0}-{1}-B".format(order_number, i+1)
                cert.save_certificate_image(filename, order_number, save_to_gdrive, '1udPWiIwHkgSJKIidTv_EzfFwNAVqeLxf')
                

                number += 1
                with open(script_dir+"/next_ord_num.txt", 'w') as f:
                    f.write(str(number))
        
        processed_order_ids.append(str(order_id))
        with open(script_dir+"/processed_order_ids.txt", 'w') as f:
            f.write('\n'.join(processed_order_ids))

        k += 1
        x = int((k+1)/len(orders)*40)
        xn = 40-x
        print("Processing new order {0}/{1}:    ".format(k, len(orders)) + "#"*x + "-"*xn)





if process_spreadsheet_orders:
    with open(script_dir+"/next_ord_num.txt", 'r') as f:
        number = int(f.read())

    with open(script_dir+"/processed_orders_spreadsheet.txt", 'r') as f:
        processed_sheet_orders = [row.strip() for row in f.read().split(',')]

    gc = gspread.service_account(filename=script_dir+"/orders-upload-326614-030cdb4cbb9d.json")
    sh = gc.open("Tracciamento ordini inviati + richieste recensioni").get_worksheet_by_id(902577955)
    sheet_list = sh.get_all_values()


    for i in range(1, len(sheet_list)):
        if not(str(i) in processed_sheet_orders) and len(sheet_list[i]) == 10 and sheet_list[i][-1] == "Yes":
            try:
                product_type = sheet_list[i][1].upper().strip()
                framesize = sheet_list[i][2].split(' ')[-1]
                date = time.strftime("%d %m %Y", time.localtime()) if len(sheet_list[i][3].strip())==0 else ' '.join(sheet_list[i][3].strip().split('/'))
                item_lang = {"IT": 0, "EN": 1}[sheet_list[i][4].upper().strip()]
                
                if len(date.split(' ')) != 3 or not (product_type in ["LORD", "LADY", "COUPLE", "FAMILY"]) or len(framesize)==0:
                    continue

                """
                "nell’anno sessantanove del"
                "nel"

                "sixty ninth year of the "
                ""
                """
                if int(date.split(' ')[2]) > 2031: continue
                try:
                    ryearspec = ["nell’anno {0} del".format(ryears[date.split(' ')[2]][item_lang]), "{0} year of the ".format(ryears[date.split(' ')[2]][item_lang])][item_lang]
                except:
                    ryearspec = ["nel", ""][item_lang]
            except:
                continue
            
            if product_type == "LORD" or product_type == "LADY":
                is_lord = product_type == "LORD"
                try:
                    name = sheet_list[i][5 if is_lord else 6].upper().strip()
                    if len(name) == 0:
                        raise
                    template_fills = [("$NAME$", name), ("$DATE$", date), ("$MQ$", mqs[framesize][item_lang]), ("$NUMBER$", str(number)), ("$MONTHDAY$", monthdays[int(date.split(' ')[0])-1][item_lang]), 
                        ("$MONTH$", months[int(date.split(' ')[1])-1][item_lang]), ("$YEAR$", years[date.split(' ')[2]][item_lang]), ("$RYEAR$", ryearspec), ("$AREA$", areas[framesize][item_lang]),
                        ("$JR$", "")]
                except:
                    continue
                
                cert = Certificate(0 if is_lord else 2, template_fills, item_lang)
                cert.make_certificate_image()
                filename = "#{0}-A".format(i)
                cert.save_certificate_image(filename, i, save_to_gdrive, '1P9gQWZ1yhmIibEoSDGfURCxV-qQit8yx')

                cert = Certificate(1 if is_lord else 3, template_fills, item_lang)
                cert.make_certificate_image()
                filename = "#{0}-B".format(i)
                cert.save_certificate_image(filename, i, save_to_gdrive, '1P9gQWZ1yhmIibEoSDGfURCxV-qQit8yx')

                number += 1
                with open(script_dir+"/next_ord_num.txt", 'w') as f:
                    f.write(str(number))
            
            elif product_type == "COUPLE":
                try:
                    name1 = sheet_list[i][5].upper().strip()
                    name2 = sheet_list[i][6].upper().strip()
                    if len(name1) == 0 or len(name2) == 0:
                        raise
                    
                    spella = ["il Lord e la Lady", "the Lord and the Lady"][item_lang]
                    spellb = ["del Lord e della Lady", "the Lord and the Lady"][item_lang]

                    template_fills = [("$TITLE1$", "LORD"), ("$NAME1$", name1), ("$TITLE2$", "LADY"), ("$NAME2$", name2), 
                                    ("$DATE$", date), ("$MQ$", mqs[framesize][item_lang]), ("$NUMBER$", str(number)), ("$MONTHDAY$", monthdays[int(date.split(' ')[0])-1][item_lang]), 
                                    ("$MONTH$", months[int(date.split(' ')[1])-1][item_lang]), ("$YEAR$", years[date.split(' ')[2]][item_lang]), ("$RYEAR$", ryearspec), 
                                    ("$AREA$", areas[framesize][item_lang]), ("$SPELLA$", spella), ("$SPELLB$", spellb)]
                except:
                    continue
                
                cert = Certificate(4, template_fills, item_lang)
                cert.make_certificate_image()
                filename = "#{0}-A".format(i)
                cert.save_certificate_image(filename, i, save_to_gdrive, '1P9gQWZ1yhmIibEoSDGfURCxV-qQit8yx')

                cert = Certificate(5, template_fills, item_lang)
                cert.make_certificate_image()
                filename = "#{0}-B".format(i)
                cert.save_certificate_image(filename, i, save_to_gdrive, '1P9gQWZ1yhmIibEoSDGfURCxV-qQit8yx')
                
                number += 1
                with open(script_dir+"/next_ord_num.txt", 'w') as f:
                    f.write(str(number))
            
            elif product_type == "FAMILY":
                members = []
                if len(sheet_list[i][5].upper().strip()) != 0:
                    members.append(("LORD", sheet_list[i][5].upper().strip()))
                if len(sheet_list[i][6].upper().strip()) != 0:
                    members.append(("LADY", sheet_list[i][6].upper().strip()))

                for lordjr in sheet_list[i][7].split(','):
                    ljr = lordjr.upper().strip()
                    if len(ljr) == 0: continue
                    members.append(("LORD", "JR. " + ljr))
                for ladyjr in sheet_list[i][8].split(','):
                    ljr = ladyjr.upper().strip()
                    if len(ljr) == 0: continue
                    members.append(("LADY", "JR. " + ljr))
                
                if len(members) < 3: continue

                for m, member in enumerate(members):
                    is_lord = member[0] == "LORD"
                    try:
                        name = member[1]
                        template_fills = [("$NAME$", name), ("$DATE$", date), ("$MQ$", mqs[framesize][item_lang]), ("$NUMBER$", str(number)), ("$MONTHDAY$", monthdays[int(date.split(' ')[0])-1][item_lang]), 
                            ("$MONTH$", months[int(date.split(' ')[1])-1][item_lang]), ("$YEAR$", years[date.split(' ')[2]][item_lang]), ("$RYEAR$", ryearspec), ("$AREA$", areas[framesize][item_lang]),
                            ("$JR$", "")]
                    except:
                        continue
                    
                    cert = Certificate(0 if is_lord else 2, template_fills, item_lang)
                    cert.make_certificate_image()
                    filename = "#{0}-{1}-A".format(i, m)
                    cert.save_certificate_image(filename, i, save_to_gdrive, '1P9gQWZ1yhmIibEoSDGfURCxV-qQit8yx')

                    cert = Certificate(1 if is_lord else 3, template_fills, item_lang)
                    cert.make_certificate_image()
                    filename = "#{0}-{1}-B".format(i, m)
                    cert.save_certificate_image(filename, i, save_to_gdrive, '1P9gQWZ1yhmIibEoSDGfURCxV-qQit8yx')

                    number += 1
                    with open(script_dir+"/next_ord_num.txt", 'w') as f:
                        f.write(str(number))


            processed_sheet_orders.append(str(i))
    
    with open(script_dir+"/processed_orders_spreadsheet.txt", 'w') as f:
        f.write(','.join(processed_sheet_orders))
