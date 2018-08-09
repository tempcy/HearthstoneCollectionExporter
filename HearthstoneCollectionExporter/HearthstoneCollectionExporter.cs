using System;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using CsvHelper;
using HearthMirror;
using HearthDb;
using HearthDb.Enums;
using System.Text;

namespace HearthstoneCollectionExporter
{
    static class HearthstoneCollectionExporter
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            var status = HearthMirror.Status.GetStatus().MirrorStatus;
            if (status == HearthMirror.Enums.MirrorStatus.Ok)
            {
                // Found Hearthstone process, attempting to get cards and write them to CSV.
                var goldenCollection = Reflection.GetCollection().Where(x => x.Premium);
                var commonCollection = Reflection.GetCollection().Where(x => !x.Premium);

                using (var textWriter = new StreamWriter("cards.csv", false, Encoding.Default))
                {
                    var csv = new CsvWriter(textWriter);
                    // Write headers
                    csv.WriteField("Id");
                    csv.WriteField("Name");
                    csv.WriteField("Card Count");
                    csv.WriteField("Need");
                    csv.WriteField("Recommend");
                    csv.WriteField("Set");
                    csv.WriteField("Zone");
                    csv.WriteField("Class");
                    csv.WriteField("Rarity");
                    csv.WriteField("Type");
                    csv.WriteField("Race");
                    csv.WriteField("Mana");
                    csv.WriteField("Attack");
                    csv.WriteField("Health");
                    csv.WriteField("Text");
                    csv.NextRecord();

                    foreach (var dbCard in Cards.Collectible)
                    {
                        csv.WriteField(dbCard.Key); // Id
                        csv.WriteField(dbCard.Value.GetLocName(Locale.zhCN));
                        var amountNormal =
                            commonCollection.Where(x => x.Id.Equals(dbCard.Key)).Select(x => x.Count).FirstOrDefault();
                        var amountGolden =
                            goldenCollection.Where(x => x.Id.Equals(dbCard.Key)).Select(x => x.Count).FirstOrDefault();
                        csv.WriteField(amountNormal + amountGolden);
                        csv.WriteField(null);
                        csv.WriteField(null);
                        csv.WriteField(dbCard.Value.Set.ToString()
                            .Replace("HOF", "荣誉室").Replace("CORE", "基本").Replace("EXPERT1", "经典")
                            .Replace("NAXX", "纳克萨玛斯").Replace("PE1", "地精大战侏儒")
                            .Replace("BRM", "黑石山的火焰").Replace("PE2", "冠军的试炼").Replace("LOE", "探险家协会")
                            .Replace("OG", "上古之神").Replace("KARA", "卡拉赞").Replace("GANGS", "加基森")
                            .Replace("UNGORO", "安戈洛").Replace("ICECROWN", "冰封王座").Replace("LOOTAPALOOZA", "狗头人")
                            .Replace("GILNEAS", "女巫森林").Replace("BOOMSDAY", "砰砰计划"));
                        csv.WriteField(null);
                        csv.WriteField(dbCard.Value.Class.ToString().Replace("DRUID", "德鲁伊").Replace("HUNTER", "猎人")
                            .Replace("MAGE", "法师").Replace("PALADIN", "圣骑士").Replace("PRIEST", "牧师")
                            .Replace("ROGUE", "潜行者").Replace("SHAMAN", "萨满祭司").Replace("WARLOCK", "术士")
                            .Replace("WARRIOR", "战士").Replace("NEUTRAL", "中立"));
                        csv.WriteField(dbCard.Value.Rarity.ToString().Replace("FREE", "基础").Replace("COMMON", "普通")
                            .Replace("RARE", "稀有").Replace("EPIC", "史诗").Replace("LEGENDARY", "传说"));
                        csv.WriteField(dbCard.Value.Type.ToString().Replace("HERO", "英雄").Replace("MINION", "随从")
                            .Replace("ABILITY", "法术").Replace("WEAPON", "武器"));
                        if (dbCard.Value.Race == HearthDb.Enums.Race.INVALID)
                        {
                            csv.WriteField(null);
                        }
                        else
                        {
                            csv.WriteField(dbCard.Value.Race.ToString().Replace("MURLOC", "鱼人").Replace("DEMON", "恶魔")
                            .Replace("MECHANICAL", "机械").Replace("ELEMENTAL", "元素").Replace("PET", "野兽")
                            .Replace("TOTEM", "图腾").Replace("PIRATE", "海盗").Replace("DRAGON", "龙")
                            .Replace("ALL", "全部"));
                        }
                        csv.WriteField(dbCard.Value.Cost); // Mana
                        csv.WriteField(dbCard.Value.Attack);
                        csv.WriteField(dbCard.Value.Health);
                        csv.WriteField(dbCard.Value.GetLocText(Locale.zhCN));
                        csv.NextRecord();
                    }

                }
                string fixtext = File.ReadAllText(@"cards.csv", Encoding.Default);
                File.WriteAllText(@"cards.csv", fixtext.Replace("$", "").Replace("<b>", "").Replace("</b>", "").Replace("<i>", "").Replace("</i>", ""), Encoding.Default);
                MessageBox.Show("Finished exporting cards to cards.csv.", "Success");
            }
            else if (status == HearthMirror.Enums.MirrorStatus.ProcNotFound)
            {
                MessageBox.Show("Unable to find Hearthstone the process.", "Error");
            }
            else if (status == HearthMirror.Enums.MirrorStatus.Error)
            {
                MessageBox.Show("There was a problem finding the Hearthstone process.", "Error");
            }
        }
    }
}
