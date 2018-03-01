package datos;

import app.Course;
import java.sql.*;
import java.util.*;

public class CoursesJDBC {

	private final String SQL_SELECT = "SELECT id,shortname,fullname FROM mdl_course";
	
	/**
     * Metodo que regresa el contenido de la tabla de cursos
     */
    public List<Course> select() {
        Connection conn = null;
        PreparedStatement stmt = null;
        ResultSet rs = null;
        Course curso = null;
        List<Course> courses = new ArrayList<Course>();
        try {
            conn = Conexion.getConnection();
            stmt = conn.prepareStatement(SQL_SELECT);
            rs = stmt.executeQuery();
            while (rs.next()) {
                int id = rs.getInt(1);
                String shortname = rs.getString(2);
                String fullname = rs.getString(3);
                /*System.out.print(" " + id_persona);
                 System.out.print(" " + nombre);
                 System.out.print(" " + apellido);
                 System.out.println();
                 */
                curso = new Course();
                curso.setId(id);
                curso.setShortname(shortname);
                curso.setFullname(fullname);
                courses.add(curso);
            }

        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            Conexion.close(rs);
            Conexion.close(stmt);
            Conexion.close(conn);
        }
        return courses;
    }
	
}
