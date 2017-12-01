package Database;

import android.content.Context;
import android.database.sqlite.SQLiteDatabase;
import android.database.sqlite.SQLiteOpenHelper;

/**
 * Created by admin on 9/14/2017.
 */

public class DatabaseHelper extends SQLiteOpenHelper {
    public static String DATABASE_NAME="EXPORTTOEXCEL";
    public static int DATABASE_VERSION=1;
    public static String TABLE_NAME="personinfo";
    public DatabaseHelper(Context context, String name, SQLiteDatabase.CursorFactory factory, int version) {
        super(context, DATABASE_NAME, factory, DATABASE_VERSION);

    }

    @Override
    public void onCreate(SQLiteDatabase sqLiteDatabase) {

        String query="create table if not exists " +TABLE_NAME+ "( NAME TEXT,ADDRESS TEXT, EMAILID TEXT, PHONENO TEXT)";
        sqLiteDatabase.execSQL(query);
    }

    @Override
    public void onUpgrade(SQLiteDatabase sqLiteDatabase, int i, int i1) {

    }
}
