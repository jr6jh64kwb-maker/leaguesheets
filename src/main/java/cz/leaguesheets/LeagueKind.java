package cz.leaguesheets;

public enum LeagueKind {
    SIX_TEAMS("6 týmů"),
    SEVEN_TEAMS("7 týmů"),
    EIGHT_TEAMS("8 týmů");

    private final String label;

    LeagueKind(String label) {
        this.label = label;
    }

    public String scriptName() {
        return switch (this) {
            case SIX_TEAMS -> "six";
            case SEVEN_TEAMS -> "seven";
            case EIGHT_TEAMS -> "eight";
        };
    }

    @Override
    public String toString() {
        return label;
    }
}
