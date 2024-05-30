import fs from 'fs';
import pdf from 'pdf-parse';
import {StarterLine} from "./StarterLine";

const filePath = "/Users/nschle85/IdeaProjects/MeldelisteParser/meldeliste.pdf";
const laufRegex = new RegExp("(Lauf) (\\d+)/(\\d+) (.*)");
const dataBuffer = fs.readFileSync(filePath);

async function getPdfText(): Promise<Array<string>|null> {
    try {
        const data = await pdf(dataBuffer);

        // console.log('Number of pages:', data.numpages);
        // console.log('Number of rendered pages:', data.numrender);
        // console.log('PDF info:', data.info);
        // console.log('PDF metadata:', data.metadata);
        // console.log('PDF version:', data.version);
        //console.log('PDF text:', data.text);
        const lines = data.text.split('\n');
        parse(lines);
        return lines;

    } catch (error) {
        console.error('Error parsing PDF:', error);
        return null;
    }
}

function toLeadingZeroString(value: string):string {
    return value.padStart(2,'0');
}
function parse(lines: Array<string>|null): void {
    let wettkampf = "";
    let lauf:string  = "";

    const parsedContent: Array<StarterLine> = [];

    lines?.forEach( line => {
        if (line.startsWith("Wettkampf")) {
            wettkampf = line;
        } else if (line.startsWith("Lauf")) {
            // Matcher laufMatcher = laufPattern.matcher(line);
            const laufMatcher = laufRegex.exec(line);
            if (laufMatcher && laufMatcher.length > 1) {
                lauf = laufMatcher[1] + " " + toLeadingZeroString(laufMatcher[2]) + "/" + toLeadingZeroString(laufMatcher[3]) + " " + laufMatcher[4];
            } else {
                console.error("No matching Lauf:" + line);
            }
        } else if (line.startsWith("Bahn")) {
            // match (Bahn), (BahnNummer), (schwimmer: Teilnehmer, Jahrgang, Verein etc), (Meldezeit)
            const r = /(Bahn) (\d+) (.+?) (\d+:\d+,\d+)/;
            // Now create matcher object.
            const m = r.exec(line);
            if (m && m.length > 1) {
                // var bahn = m.group(1);
                const bahnNr = m[2];
                let schwimmer = m[3];

                //extract verein if possible and update schwimmer and verein
                let verein = "";

                // Phdksfh, Fdsds  2013/ TSV Neuburg
                const schwimmerVereinPattern = /(.+?)\/ (.*)/;
                const schwimmerVereinMatcher = schwimmerVereinPattern.exec(schwimmer);

                // Rdsds, Sfdfd  1978/AK 45	SV Lohhof
                const schwimmerAkVereinPattern = /(.+?\/AK \d+) (.*)/;
                const schwimmerAkVereinMatcher = schwimmerAkVereinPattern.exec(schwimmer);

                //Rdsds, Sfdfd  1978 45	SV Lohhof
                const schwimmerJgVereinPattern = /(.+? \d+) (.*)/;
                const schwimmerJgVereinMatcher = schwimmerJgVereinPattern.exec(schwimmer);

                //Rdsds, Sfdfd  Offen 45	SV Lohhof
                const schwimmerOffenVereinPattern = /(.+? Offen) (.*)/;
                const schwimmerOffenVereinMatcher = schwimmerOffenVereinPattern.exec(schwimmer);

                if (schwimmerVereinMatcher && schwimmerVereinMatcher.length>1) {
                         schwimmer = schwimmerVereinMatcher[1];
                         verein = schwimmerVereinMatcher[2];
                }
                else if (schwimmerAkVereinMatcher && schwimmerAkVereinMatcher.length>1) {
                         schwimmer = schwimmerAkVereinMatcher[1];
                         verein = schwimmerAkVereinMatcher[2];
                }
                else if (schwimmerJgVereinMatcher && schwimmerJgVereinMatcher.length>1) {
                         schwimmer = schwimmerJgVereinMatcher[1];
                         verein = schwimmerJgVereinMatcher[2];
                }else if (schwimmerOffenVereinMatcher && schwimmerOffenVereinMatcher.length>1) {
                         schwimmer = schwimmerOffenVereinMatcher[1];
                         verein = schwimmerOffenVereinMatcher[2];
                }

                else {
                        console.error("Could not parse: ",schwimmer);
                }
                const meldezeit = m[4];
                parsedContent.push(new StarterLine(wettkampf, lauf, bahnNr, schwimmer, verein, meldezeit));
            }
        } else {
            console.error("NO MATCH ", line);
        }
    });
    console.log(parsedContent);
}

console.log('hello');
getPdfText().then(result => {parse(result)}).then( () => console.log('done'));

