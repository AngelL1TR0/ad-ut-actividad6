package entity;

public class ComicInfoEntity {
    private String superHeroe;
    private String company;
    private String creator;
    private String firstApparition;
    private String dateApparition;

    public ComicInfoEntity(String superHeroe, String company, String creator, String firstApparition, String dateApparition) {
        this.superHeroe = superHeroe;
        this.company = company;
        this.creator = creator;
        this.firstApparition = firstApparition;
        this.dateApparition = dateApparition;
    }

    public String getSuperHeroe() {
        return superHeroe;
    }

    public void setSuperHeroe(String superHeroe) {
        this.superHeroe = superHeroe;
    }

    public String getCompany() {
        return company;
    }

    public void setCompany(String company) {
        this.company = company;
    }

    public String getCreator() {
        return creator;
    }

    public void setCreator(String creator) {
        this.creator = creator;
    }

    public String getFirstApparition() {
        return firstApparition;
    }

    public void setFirstApparition(String firstApparition) {
        this.firstApparition = firstApparition;
    }

    public String getDateApparition() {
        return dateApparition;
    }

    public void setDateApparition(String dateApparition) {
        this.dateApparition = dateApparition;
    }
}
