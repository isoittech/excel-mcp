package jp.isoittech;

import com.google.gson.annotations.SerializedName;
import java.util.List;

public class JsonData {
    @SerializedName("表紙")
    private List<CellInfo> coverPage;

    @SerializedName("明細")
    private List<CellInfo> details;

    @SerializedName("表紙 (バイヤ)")
    private List<CellInfo> coverPageBuyer;

    @SerializedName("明細 (バイヤ)")
    private List<CellInfo> detailsBuyer;

    public JsonData() {}

    public List<CellInfo> getCoverPage() {
        return coverPage;
    }

    public List<CellInfo> getDetails() {
        return details;
    }

    public List<CellInfo> getCoverPageBuyer() {
        return coverPageBuyer;
    }

    public List<CellInfo> getDetailsBuyer() {
        return detailsBuyer;
    }

    @Override
    public String toString() {
        return "ReportData{" +
               "coverPage=" + coverPage +
               ", details=" + details +
               ", coverPageBuyer=" + coverPageBuyer +
               ", detailsBuyer=" + detailsBuyer +
               '}';
    }
}
