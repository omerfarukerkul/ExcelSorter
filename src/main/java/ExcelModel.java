import java.util.Objects;

public class ExcelModel {
    private String c0;
    private String c1;
    private Integer c2;

    public String getC0() {
        return c0;
    }

    public void setC0(String c0) {
        this.c0 = c0;
    }

    public String getC1() {
        return c1;
    }

    public void setC1(String c1) {
        this.c1 = c1;
    }

    public Integer getC2() {
        return c2;
    }

    public void setC2(Integer c2) {
        this.c2 = c2;
    }

    public ExcelModel(String c0, String c1, Integer c2) {
        this.c0 = c0;
        this.c1 = c1;
        this.c2 = c2;
    }

    public ExcelModel() {
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        ExcelModel that = (ExcelModel) o;
        return Objects.equals(c0, that.c0) &&
                Objects.equals(c1, that.c1) &&
                Objects.equals(c2, that.c2);
    }

    @Override
    public int hashCode() {
        return Objects.hash(c0, c1, c2);
    }

    @Override
    public String toString() {
        return "ExcelModel{" +
                "c0='" + c0 + '\'' +
                ", c1='" + c1 + '\'' +
                ", c2=" + c2 +
                '}';
    }
}
