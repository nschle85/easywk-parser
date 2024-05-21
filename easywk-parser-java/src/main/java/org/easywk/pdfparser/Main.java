package org.easywk.pdfparser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    static String wettkampf = "";
    static String lauf = "";

    private static String toLeadingZeroString(String value) {
        return String.format("%02d",Integer.parseInt(value));
    }

    public static void main(String[] args) {
        String filePath = "/Users/nschle85/IdeaProjects/MeldelisteParser/meldeliste.pdf";
        try {
            // Öffne das PDF-Dokument
            PDDocument document = PDDocument.load(new File(filePath));
            // Initialisiere PDFTextStripper
            PDFTextStripper pdfStripper = new PDFTextStripper();
            // Extrahiere den Text aus dem Dokument
            String text = pdfStripper.getText(document);
            // Schließe das Dokument
            document.close();
            // Gib den extrahierten Text aus
            List<JTeilnehmerZeile> parsedContent = new ArrayList<>();
            BufferedReader bufReader = new BufferedReader(new StringReader(text));
            String line;

            String laufRegex = "(Lauf) (\\d+)/(\\d+) (.*)";
            Pattern laufPattern = Pattern.compile(laufRegex);

            while( (line=bufReader.readLine()) != null ) {
                if (line.startsWith("Wettkampf")) {
                    wettkampf = line;
                }
                else if (line.startsWith("Lauf")) {
                    Matcher laufMatcher = laufPattern.matcher(line);
                    if(laufMatcher.find()) {
                      // Lauf 1/2 (ca xyz)
                      lauf = laufMatcher.group(1)+" "+toLeadingZeroString(laufMatcher.group(2))+"/"+toLeadingZeroString(laufMatcher.group(3))+" "+ laufMatcher.group(4);
                    }
                    else {
                        System.err.println("No matching Lauf:"+ line);
                    }
                }
                else if (line.startsWith("Bahn")){

                    // match (Bahn), (BahnNummer), (schwimmer: Teilnehmer, Jahrgang, Verein etc), (Meldezeit)
                    String pattern = "(Bahn) (\\d+) (.+?) (\\d+:\\d+,\\d+)";
                    // Create a Pattern object
                    Pattern r = Pattern.compile(pattern);
                    // Now create matcher object.
                    Matcher m = r.matcher(line);
                    if (m.find()) {
                        // var bahn = m.group(1);
                        var bahnNr = m.group(2);
                        var schwimmer = m.group(3);

                        //extract verein if possible and update schwimmer and verein
                        var verein = "";

                        //Phdksfh, Fdsds  2013/ TSV Neuburg
                        Pattern schwimmerVereinPattern = Pattern.compile("(.+?)/ (.*)");
                        Matcher schwimmerVereinMatcher = schwimmerVereinPattern.matcher(schwimmer);

                        //Rdsds, Sfdfd  1978/AK 45	SV Lohhof
                        Pattern schwimmerAkVereinPattern = Pattern.compile("(.+?/AK \\d+) (.*)");
                        Matcher schwimmerAkVereinMatcher = schwimmerAkVereinPattern.matcher(schwimmer);

                        //Rdsds, Sfdfd  1978 45	SV Lohhof
                        Pattern schwimmerJgVereinPattern = Pattern.compile("(.+? \\d+) (.*)");
                        Matcher schwimmerJgVereinMatcher = schwimmerJgVereinPattern.matcher(schwimmer);

                        //Rdsds, Sfdfd  Offen 45	SV Lohhof
                        Pattern schwimmerOffenVereinPattern = Pattern.compile("(.+? Offen) (.*)");
                        Matcher schwimmerOffenVereinMatcher = schwimmerOffenVereinPattern.matcher(schwimmer);

                        if (schwimmerVereinMatcher.find()) {
                            schwimmer = schwimmerVereinMatcher.group(1);
                            verein = schwimmerVereinMatcher.group(2);
                        }else if (schwimmerAkVereinMatcher.find()) {
                            schwimmer = schwimmerAkVereinMatcher.group(1);
                            verein = schwimmerAkVereinMatcher.group(2);
                        }else if (schwimmerJgVereinMatcher.find()) {
                            schwimmer = schwimmerJgVereinMatcher.group(1);
                            verein = schwimmerJgVereinMatcher.group(2);
                        }else if (schwimmerOffenVereinMatcher.find()) {
                            schwimmer = schwimmerOffenVereinMatcher.group(1);
                            verein = schwimmerOffenVereinMatcher.group(2);
                        }
                        else {
                            System.err.println("Could not parse: "+schwimmer);
                        }
                        var meldezeit = m.group(4);
                        parsedContent.add(new JTeilnehmerZeile(wettkampf, lauf, bahnNr,  schwimmer, verein, meldezeit));
                    } else {
                        System.err.println("NO MATCH "+ line);
                    }
                }
                else {
                }

                File fnew = new File(filePath+".csv");
                FileWriter f2;
                try {
                    f2 = new FileWriter(fnew,false);
                    f2.write("Wettkampf\tLauf\tBahn\tSchwimmer\tVerein\tMeldezeit\n");
                    parsedContent.forEach(jTeilnehmerZeile -> {
                        try {
                            System.out.println(jTeilnehmerZeile);
                            f2.write(jTeilnehmerZeile+"\n");
                        } catch (IOException e) {
                            throw new RuntimeException(e);
                        }
                    });
                    f2.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}