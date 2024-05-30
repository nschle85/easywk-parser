export class StarterLine {
    public readonly wettkampf: string;
    public readonly lauf: string;
    public readonly bahn: string;
    public readonly teilnehmer: string;
    public readonly verein: string;
    public readonly meldezeit: string;


    constructor(wettkampf: string, lauf:string, bahn:string, teilnehmer: string, verein: string, meldezeit: string) {
        this.wettkampf = wettkampf;
        this.lauf = lauf;
        this.bahn = bahn;
        this.teilnehmer = teilnehmer;
        this.verein = verein;
        this.meldezeit = meldezeit;
    }

    public toString(): string {
        return this.wettkampf+"\t"+this.lauf+"\t"+this.bahn+"\t"+this.teilnehmer+"\t"+this.verein+"\t"+this.meldezeit;
    }
}