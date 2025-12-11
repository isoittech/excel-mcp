package jp.isoittech;

public class CellInfo {
    private int row;
    private int column;
    private String type;
    // valueの型はJSONによって異なる可能性があるのでObject型
    private Object value;

    public CellInfo() {}

    // Getterメソッド
    public int getRow() { return row; }
    public int getColumn() { return column; }
    public String getType() { return type; }
    public Object getValue() { return value; } 

    public String getStringValue() {
        return (String)this.value;
    }

    public int getIntValue() {
        return (int)((java.lang.Double)this.value).doubleValue();
    }

    public double getDoubleValue() {
        return (double)this.value;
    }

    @Override
    public String toString() {
        String exp = null;
        switch (getType()) {
        case "string":
            exp = "CellInfo{" +
                "row=" + row +
                ", column=" + column +
                ", type='" + type + '\'' +
                ", value=" + getStringValue() + 
                '}';
            break;
        case "integer":
            exp = "CellInfo{" +
                "row=" + row +
                ", column=" + column +
                ", type='" + type + '\'' +
                ", value=" + getIntValue() +
                '}';
            break;
        case "double":
            exp = "CellInfo{" +
                "row=" + row +
                ", column=" + column +
                ", type='" + type + '\'' +
                ", value=" + getDoubleValue() + 
                '}';
            break;

        }
        return exp;
    }
}

