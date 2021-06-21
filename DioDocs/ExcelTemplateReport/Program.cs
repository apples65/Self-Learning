using GrapeCity.Documents.Excel;
using System.IO;
using System.Text.Json;

namespace DioDocsExcelReport2
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Hello DioDocs!");

            // データソース（JSON）
            // var jsonString = File.ReadAllText("Data/testdata_multicustomer.json");
            // var data = JsonSerializer.Deserialize<Data>(jsonString);


            // データソース（カスタムオブジェクト）
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
                            new Detail{ sku = "01-105", name = "モッツァレッラチーズとトマトのサラダ", price = 1200, unit = 70, remark = "前菜" },
                            new Detail{ sku = "02-107", name = "娼婦風スパゲティー", price = 2500, unit = 100, remark = "パスタ" },
                            new Detail{ sku = "03-120", name = "小羊背肉のリンゴソースかけ", price = 5000, unit = 130, remark = "メイン" },
                            new Detail{ sku = "04-101", name = "ごま蜜団子", price = 800, unit = 60, remark = "デザート" }
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
                            new Detail{ sku = "02-206", name = "激辛麻婆焼きそば", price = 1000, unit = 50, remark = "麺" },
                            new Detail{ sku = "03-208", name = "塩竈マグロ", price = 2300, unit = 80, remark = "魚" },
                            new Detail{ sku = "04-230", name = "超極厚牛タン", price = 4800, unit = 110, remark = "肉" },
                            new Detail{ sku = "05-202", name = "濃厚ずんだロール", price = 800, unit = 60, remark = "デザート" },
                            new Detail{ sku = "06-227", name = "ハチミツパンケーキ", price = 1300, unit = 60, remark = "お持ち帰り" }
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
                            new Detail{ sku = "03-307", name = "マスクメロン", price = 1100, unit = 60, remark = "フルーツ" },
                            new Detail{ sku = "04-309", name = "シャインマスカット", price = 2400, unit = 90, remark = "フルーツ" },
                            new Detail{ sku = "05-340", name = "ピオーネ", price = 4900, unit = 120, remark = "フルーツ" },
                            new Detail{ sku = "06-303", name = "ラ・フランス", price = 700, unit = 50, remark = "フルーツ" },
                            new Detail{ sku = "07-338", name = "富有柿", price = 1400, unit = 70, remark = "フルーツ" },
                            new Detail{ sku = "05-116", name = "サンジェルマンのサンドイッチバッグ", price = 1500, unit = 80, remark = "お持ち帰り" }
                        }
                    }
                }
            };

            // ライセンスキー
            // string key = "hogehoge";
            // Workbook.SetLicenseKey(key);

            // 新しいワークブックを生成
            var workbook = new Workbook();

            // テンプレートを読み込む
            workbook.Open("Templates/SimpleInvoiceJP_可変明細_単一シート.xlsx");
            // workbook.Open("Templates/SimpleInvoiceJP_可変明細_複数シート.xlsx");
            // workbook.Open("Templates/SimpleInvoiceJP_可変明細_複数シート_数式維持.xlsx");
            // workbook.Open("Templates/SimpleInvoiceJP_固定明細_単一シート.xlsx");
            // workbook.Open("Templates/SimpleInvoiceJP_固定明細_複数シート.xlsx");
            // workbook.Open("Templates/SimpleInvoiceJP_固定明細_複数シート_数式維持.xlsx");

            // データソースを追加
            workbook.AddDataSource("ds", data);

            // テンプレート処理を呼び出し
            workbook.ProcessTemplate();

            // Excelファイルに保存
            workbook.Save("result.xlsx");

            // PDFファイルに保存
            workbook.Save("result.pdf", SaveFileFormat.Pdf);
        }
    }
}
