/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package read_write;

/**
 *
 * @author S530671
 */
public class Album {

    int SNO;
    String genre;
    int criticScore;
    String albumName;
    String artist;
    String date;

    public Album(int SNO, String genre, int criticScore, String albumName, String artist, String date) {
        this.SNO = SNO;
        this.genre = genre;
        this.criticScore = criticScore;
        this.albumName = albumName;
        this.artist = artist;
        this.date = date;
    }

    @Override
    public String toString() {
        return SNO + " " + genre +" " +  criticScore +" " + albumName + " " +  
                artist + " " + date;
    }

}
