/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package absensi.psi;

/**
 *
 * @author Lab04
 */
public class Employee {
    private String id;
    private String date;
    private String time;
    private String dateTime;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getTime() {
        return time;
    }

    public void setTime(String time) {
        this.time = time;
    }
    public void mergeDateAdnTime(){
        dateTime=date+time;
    }

    public String getDateTime() {
        return dateTime;
    }   

    @Override
    public String toString() {
        return "Employee {" + "id=" + id + ", date=" +date+time+ ")";
    }
    
    
    
    
     
}
