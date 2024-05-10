package org.example;

public class JTeilnehmerZeile
{
    public final String wettkampf;
    public final String lauf;
    public final String bahn;
    public final String teilnehmer;
    public final String verein;
    public final String meldezeit;


    JTeilnehmerZeile(String wettkampf, String lauf, String bahn, String teilnehmer, String verein, String meldezeit)
    {
        this.wettkampf = wettkampf;
        this.lauf = lauf;
        this.bahn = bahn;
        this.teilnehmer = teilnehmer;
        this.verein = verein;
        this.meldezeit = meldezeit;
    }

    @Override
    public String toString() {
        return this.wettkampf+"\t"+this.lauf+"\t"+bahn+"\t"+this.teilnehmer+"\t"+this.verein+"\t"+this.meldezeit;
    }
}

