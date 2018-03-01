package app;

public class Course {

	private String shortname;
	private String fullname;
	private int    id;

	public String getShortname() {
		return shortname;
	}

	public void setShortname(String shortname) {
		this.shortname = shortname;
	}


	public String getFullname() {
		return fullname;
	}

	public void setFullname(String fullname) {
		this.fullname = fullname;
	}

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}


    /* @Override
    public String toString() {
        return "Curso{" + "nombre_corto=" + shortname + ", nombre_curso=" + fullname + '}';
     }*/


}
