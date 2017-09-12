package com.liqiaoqi.util;

import java.util.ArrayList;

/**
 * Created by LQQ on 2017/9/12 0012.
 */
public class TyphoonInfluence {
    private String year; //年份
    private int num_of_typhoon; //该年台风总数
    private int num_of_nb_typhoon;  //影响宁波台风数
    private String proportion; //比例
    private ArrayList<String> typhooncodes;//具体台风码

    public TyphoonInfluence(String year,int num_of_typhoon){
        this.year = year;
        this.num_of_typhoon = 1;
        this.typhooncodes = new ArrayList<String>();
    }

    public String getYear() {
        return year;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public int getNum_of_typhoon() {
        return num_of_typhoon;
    }

    public void setNum_of_typhoon(int num_of_typhoon) {
        this.num_of_typhoon = num_of_typhoon;
    }

    public int getNum_of_nb_typhoon() {
        return num_of_nb_typhoon;
    }

    public void setNum_of_nb_typhoon(int num_of_nb_typhoon) {
        this.num_of_nb_typhoon = num_of_nb_typhoon;
    }

    public String getProportion() {
        return proportion;
    }

    public void setProportion(String proportion) {
        this.proportion = proportion;
    }

    public ArrayList<String> getTyphooncodes() {
        return typhooncodes;
    }

    public void setTyphooncodes(ArrayList<String> typhooncodes) {
        this.typhooncodes = typhooncodes;
    }
}
