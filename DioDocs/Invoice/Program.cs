using GrapeCity.Documents.Excel;

// See https://aka.ms/new-console-template for more information
Console.WriteLine("Hello, World!");

var data = new Data
{
    publisher = new Publisher
    {
        companyname = "ディオドック株式会社",
        postalcode = "981-2222",
        address1 = "M県S市紅葉区",
        address2 = "杜王町2-6-11",
        tel = "022-777-8210",
        bankname = "むかでや銀行",
        bankblanch = "杜王町支店",
        account = "123-456789",
        representative = "葡萄　城太郎"
    },
    customer = new Customer[]
    {
        new Customer
        {
            companyname = "財団法人スピードワゴン",
            name = "虎猿出井　富似雄",
            postalcode = "981-9999",
            address1 = "M県S市広瀬区",
            address2 = "花京院3-1-4",
            tel = "022-987-2220",
            detail = new Detail[]
            {
                new Detail{ sku = "0105231", name = "モッツァレッラチーズとトマトのサラダ", price = 1200, unit = 70, remark = "前菜" },
                new Detail{ sku = "0201071", name = "娼婦風スパゲティー", price = 2500, unit = 100, remark = "パスタ" },
                new Detail{ sku = "0301201", name = "小羊背肉のリンゴソースかけ", price = 5000, unit = 130, remark = "メイン" },
                new Detail{ sku = "0401011", name = "ごま蜜団子", price = 800, unit = 60, remark = "デザート" }
            }
        },
        new Customer
        {
            companyname = "杜王グランドホテル",
            name = "江陽　太郎",
            postalcode = "982-1234",
            address1 = "M県S市秋保区",
            address2 = "上愛子6-4-7",
            tel = "022-876-1119",
            detail = new Detail[]
            {
                new Detail{ sku = "0202061", name = "激辛麻婆焼きそば", price = 1000, unit = 50, remark = "麺" },
                new Detail{ sku = "0302081", name = "塩竈マグロ", price = 2300, unit = 80, remark = "魚" },
                new Detail{ sku = "0402301", name = "超極厚牛タン", price = 4800, unit = 110, remark = "肉" },
                new Detail{ sku = "0502021", name = "濃厚ずんだロール", price = 800, unit = 60, remark = "デザート" },
                new Detail{ sku = "0602271", name = "ハチミツパンケーキ", price = 1300, unit = 60, remark = "お持ち帰り" }
            }
        },
        new Customer
        {
            companyname = "東方フルーツパーラー",
            name = "東方　花子",
            postalcode = "983-5678",
            address1 = "M県S市荒井区",
            address2 = "深沼9-7-1",
            tel = "022-765-0008",
            detail = new Detail[]
            {
                new Detail{ sku = "0303071", name = "マスクメロン", price = 1100, unit = 60, remark = "フルーツ" },
                new Detail{ sku = "0403091", name = "シャインマスカット", price = 2400, unit = 90, remark = "フルーツ" },
                new Detail{ sku = "0503401", name = "ピオーネ", price = 4900, unit = 120, remark = "フルーツ" },
                new Detail{ sku = "0603031", name = "ラ・フランス", price = 700, unit = 50, remark = "フルーツ" },
                new Detail{ sku = "0703381", name = "富有柿", price = 1400, unit = 70, remark = "フルーツ" },
                new Detail{ sku = "0501161", name = "サンジェルマンのサンドイッチバッグ", price = 1500, unit = 80, remark = "お持ち帰り" }
            }
        }
    }
};

// 新しいワークブックを生成
var workbook = new Workbook();

// テンプレートを読み込む
// workbook.Open("Templates/請求書テンプレート.xlsx");
// workbook.Open("Templates/請求書テンプレート（可変明細）.xlsx");
// workbook.Open("Templates/請求書テンプレート（単一シート）.xlsx");
workbook.Open("Templates/請求書テンプレート（備考欄とバーコード）.xlsx");

// データソースを追加
workbook.AddDataSource("ds", data);

// テンプレート処理を呼び出し
workbook.ProcessTemplate();

// Excelファイルに保存
workbook.Save("result.xlsx");

// PDFファイルに保存
workbook.Save("result.pdf");
