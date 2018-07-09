/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package read_write;

import java.util.Comparator;

/**
 *
 * @author S530671
 */
public class AlbumSort implements Comparator<Album>
{

    @Override
    public int compare(Album a1, Album a2) {
        return a2.criticScore-a1.criticScore;
    }
    
}
