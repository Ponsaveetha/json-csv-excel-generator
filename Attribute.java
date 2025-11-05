package com.example.demo.model;

public class Attribute {

    private String name;
    private String type;

    public Attribute() {
    }

    public Attribute(String name, String type) {
        this.name = name;
        this.type = type;
    }

    // 🔹 Getters and Setters
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    @Override
    public String toString() {
        return "Attribute{" +
                "name='" + name + '\'' +
                ", type='" + type + '\'' +
                '}';
    }
}
