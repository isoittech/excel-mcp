package msi.aiproc.excel;

import com.google.gson.Gson;
import com.google.gson.JsonSyntaxException;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.List;

public class ReadJsonFile {

    public static JsonData read(String jsonFilePath) {
        JsonData jsonData = null;
        try (FileReader reader = new FileReader(jsonFilePath)) {

            // Gsonオブジェクトを作成
            Gson gson = new Gson();

            // JSONファイルをJsonDataオブジェクトに直接マッピングして読み込む
            jsonData = gson.fromJson(reader, JsonData.class);

        } catch (FileNotFoundException e) {
            System.err.println("指定されたファイルが見つかりません: " + jsonFilePath);
            e.printStackTrace();
        } catch (IOException e) {
            System.err.println("ファイルの読み込み中にエラーが発生しました:");
            e.printStackTrace();
        } catch (JsonSyntaxException e) {
            System.err.println("JSONの構文エラーまたはマッピングエラーです。JSON構造とJavaクラス定義を確認してください:");
            e.printStackTrace();
        } catch (Exception e) {
            System.err.println("処理中に予期せぬエラーが発生しました:");
            e.printStackTrace();
        }
        return jsonData;
    }
}
