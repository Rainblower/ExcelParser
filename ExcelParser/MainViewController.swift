//
//  MainViewController.swift
//  ExcelParser
//
//  Created by Admin on 22.06.2019.
//  Copyright © 2019 Admin. All rights reserved.
//

import Cocoa

import CSV
import Alamofire
import SwiftyJSON
import CoreXLSX

class MainViewController: NSViewController {

    @IBOutlet var textView: NSTextView!
    @IBOutlet var consoleView: NSTextView!
    
    @IBAction func selectFile(_ sender: Any) {
        
        let dialog = NSOpenPanel()
        dialog.title                   = "Select .xlsx file"
        dialog.showsResizeIndicator    = true;
        dialog.showsHiddenFiles        = false;
        dialog.canChooseDirectories    = true;
        dialog.canCreateDirectories    = true;
        dialog.allowsMultipleSelection = false;
        dialog.allowedFileTypes        = ["xlsx"];
        
        if dialog.runModal() == .OK {
            if dialog.url != nil {
                xlsxToHtml(url: dialog.url!)
            } else {
                return
            }
        }
    }
    
    func xlsxToHtml(url: URL) {
       
        var collumnCount = 0
        var rowCount = 0
        var count = 0

        guard let file = XLSXFile(filepath: url) else {
            fatalError("XLSX file corrupted or does not exist")
        }
        
        do {
            let path = try file.parseWorksheetPaths()
            let ws = try file.parseWorksheet(at: path[2])
            for row in ws.data?.rows ?? [] {
                rowCount += 1
                
                if rowCount < 4 {
                    continue
                }
                
                if row.cells.count == 6 {
                    continue
                }
                
                if row.cells.count == 8 {
                    var nilCount = 0
                    for cell in row.cells {
                        if cell.value == nil {
                            nilCount += 1
                        }
                    }
                    
                    if nilCount == 8 {
                        continue
                    }
                }
                
                for cell in row.cells {

                    if collumnCount < 4 {
                     collumnCount += 1
                        continue
                    }
                    
                    if row.cells.count > 8 {
                        if cell == row.cells[8]{
                            continue
                        }
                    }
                   
                    count += 1
                    let value = String(cell.value ?? "")
                    html = html.replacingOccurrences(of: ">$\(count)<", with: ">\(value)<")
                    print(value)
                }
                print("Row end")
                 collumnCount = 0
            }
            
        } catch {}
        
        textView.string = html
        updatePage(Date())
        updatePageField(html)
    }
    
    func updatePage(_ dateNow: Date) {
        guard let url = URL(string: "http://www.kp11.ru/api/page/update.php") else { return }
        
        let dateFormatter = DateFormatter()
        dateFormatter.dateFormat = "dd.MM.yy"
        
        let modDateFormatter = DateFormatter()
        modDateFormatter.dateFormat = "yy-MM-dd HH:mm:ss"
        
        let date = dateFormatter.string(from: dateNow)
        let modDate = modDateFormatter.string(from: dateNow)
        
        let params: Parameters = [
            "id" : "72",
            "page_type_id" : "14",
            "title" : "Количество поданных заявлений",
            "page_h1" : "Количество поданных заявлений на \(date)",
            "parent" : "65",
            "url" : "/abiturientu/kolichestvo_podannyh_zayavlenij",
            "active" : "1",
            "create_date" : "2015-02-06 00:17:32",
            "modify_date" : "\(modDate)"
        ]
        
        Alamofire.request(url, method: .post, parameters: params, encoding: JSONEncoding.default)
            .validate()
            .responseJSON { (response) in
                switch response.result {
                case .success(let value):
                    let json = JSON(value)
//                    self.printConsole(json.stringValue)
                    print(value)
                case .failure(let error):
                    print(error)
                }
        }
    }
    
    func updatePageField(_ html: String) {
        
        guard let url = URL(string: "http://www.kp11.ru/api/page_field/update.php") else { return }
        let params: Parameters = [
            "id" : "72",
            "page_text" : html,
            "page_right_text" : "",
            "keywords" : "заявление, количество поданных заявлений, подать заявление"
        ]
        
   
        
        Alamofire.request(url, method: .post, parameters: params, encoding: JSONEncoding.default)
            .validate()
            .responseJSON { (response) in
                switch response.result {
                case .success(let value):
                    let json = JSON(value)
                    self.printConsole(json.stringValue)
                    print(value)
                case .failure(let error):
//                    printConsole(error)
                    print(error)
                }
        }
    }
    
    func printConsole(_ message: String) {
        consoleView.string = consoleView.string + "\n" + message
        print(message)
    }
    
    
    var html = """
        <table>
        <tbody>
        <tr style="height: 108px;">
        <td class="xl64" style="width: 200px; height: 108px;"><strong>Специальность</strong></td>
        <td class="xl64" style="width: 109px; height: 108px;"><strong>Код специальности</strong></td>
        <td class="xl64" style="width: 69px; height: 108px;"><strong>Форма обучения</strong></td>
        <td class="xl64" style="width: 200px; height: 108px;"><strong>База приема</strong></td>
        <td class="xl72" style="width: 10px; height: 108px;"><strong>План набора (бюджет)</strong></td>
        <td class="xl72" style="width: 79px; height: 108px;"><strong>План</strong><br /><strong>набора (внебюджет)</strong></td>
        <td class="xl72" style="width: 77px; height: 108px;"><strong>Подано заявлений (бюджет)</strong></td>
        <td class="xl72" style="width: 94px; height: 108px;"><strong>Подано заявлений (внебюджет)</strong></td>
        </tr>
        <tr style="height: 20px;">
        <td class="xl77" style="font-size: 15pt; font-family: Bold; vertical-align: middle; text-align: center;height: 20px; width: 1000px;" colspan="8"><strong>Центр медицинской техники и оптики</strong></td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl65" style="width: 200px; height: 18px;">Медицинская оптика</td>
        <td class="xl65" style="width: 109px; height: 18px;">31.02.04</td>
        <td class="xl65" style="width: 69px; height: 18px;">очная</td>
        <td class="xl65" style="width: 175px; height: 18px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$1</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$2</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$3</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$4</td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl65" style="width: 200px; height: 18px;">Медицинская оптика</td>
        <td class="xl65" style="width: 109px; height: 18px;">31.02.04</td>
        <td class="xl65" style="width: 69px; height: 18px;">очная</td>
        <td class="xl65" style="width: 175px; height: 18px;">11 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$5</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$6</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$7</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$8</td>
        </tr>
        <tr style="height: 90px;">
        <td class="xl65" style="width: 200px; height: 90px;">Монтаж, техническое обслуживание и ремонт медицинской техники</td>
        <td class="xl76" style="width: 109px; height: 90px;">12/02/07</td>
        <td class="xl65" style="width: 69px; height: 90px;">очная</td>
        <td class="xl65" style="width: 175px; height: 90px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$9</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$10</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$11</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$12</td>
        </tr>
        <tr style="height: 36px;">
        <td class="xl65" style="width: 200px; height: 36px;">Аддитивные технологии</td>
        <td class="xl76" style="width: 109px; height: 36px;">15/02/09</td>
        <td class="xl65" style="width: 69px; height: 36px;">очная</td>
        <td class="xl65" style="width: 175px; height: 36px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$13</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$14</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$15</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$16</td>
        </tr>
        <tr style="height: 36px;">
        <td class="xl65" style="width: 200px; height: 36px;">Медицинская оптика (УП)</td>
        <td class="xl65" style="width: 109px; height: 36px;">31.02.04</td>
        <td class="xl65" style="width: 69px; height: 36px;">заочная</td>
        <td class="xl65" style="width: 175px; height: 36px;">11 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$17</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$18</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$19</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$20</td>
        </tr>
        <tr style="height: 36px;">
        <td class="xl65" style="width: 200px; height: 36px;">
        <p class="p1"><span class="s1">Медицинская оптика (БП)</span></p>
        </td>
        <td class="xl65" style="width: 200px; height: 36px;">31.02.04</td>
        <td class="xl65" style="width: 69px; height: 36px;">заочная</td>
        <td class="xl65" style="width: 175px; height: 36px;">11 кл.&nbsp;</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$21</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$22</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$23</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$24</td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl66" style="width: 200px; height: 18px; text-align: right;" colspan="4"><strong>Итого:</strong></td>
        <td class="xl72" style="width: 10px; height: 18px; text-align: right;"><strong>$25</strong></td>
        <td class="xl72" style="width: 79px; height: 18px; text-align: right;"><strong>$26</strong></td>
        <td class="xl72" style="width: 77px; height: 18px; text-align: right;"><strong>$27</strong></td>
        <td class="xl72" style="width: 94px; height: 18px; text-align: right;"><strong>$28</strong></td>
        </tr>
        <tr style="height: 20px;">
        <td class="xl77" style="font-size: 15pt; font-family: Bold; vertical-align: middle; text-align: center;vheight: 20px; width: 1000px;" colspan="8"><strong>Центр информационно-коммуникационных технологий</strong></td>
        </tr>
        <tr style="height: 54px;">
        <td class="xl65" style="width: 200px; height: 54px;">Информационные системы и программирование</td>
        <td class="xl76" style="width: 109px; height: 54px;">09/02/07</td>
        <td class="xl65" style="width: 69px; height: 54px;">очная</td>
        <td class="xl65" style="width: 175px; height: 54px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$29</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$30</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$31</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$32</td>
        </tr>
        <tr style="height: 54px;">
        <td class="xl65" style="width: 200px; height: 54px;">Информационные системы и программирование</td>
        <td class="xl76" style="width: 109px; height: 54px;">09/02/07</td>
        <td class="xl65" style="width: 69px; height: 54px;">очная</td>
        <td class="xl65" style="width: 175px; height: 54px;">11 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$33</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$34</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$35</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$36</td>
        </tr>
        <tr style="height: 54px;">
        <td class="xl65" style="width: 200px; height: 54px;">Обеспечение информационной безопасности автоматизированных систем</td>
        <td class="xl76" style="width: 109px; height: 54px;">10/02/05</td>
        <td class="xl65" style="width: 69px; height: 54px;">очная</td>
        <td class="xl65" style="width: 175px; height: 54px;"><p>11 кл.</p></td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$37</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$38</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$39</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$40</td>
        </tr>
        <tr style="height: 36px;">
        <td class="xl65" style="width: 200px; height: 36px;">Сетевое и системное администрирование</td>
        <td class="xl76" style="width: 109px; height: 36px;">09/02/06</td>
        <td class="xl65" style="width: 69px; height: 36px;">очная</td>
        <td class="xl65" style="width: 175px; height: 36px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$41</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$42</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$43</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$44</td>
        </tr>
        <tr style="height: 54px;">
        <td class="xl65" style="width: 200px; height: 54px;">Компьютерные системы и комплексы</td>
        <td class="xl76" style="width: 109px; height: 54px;">09/02/01</td>
        <td class="xl65" style="width: 69px; height: 54px;">очная</td>
        <td class="xl65" style="width: 175px; height: 54px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$45</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$46</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$47</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$48</td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl66" style="width: 200px; height: 18px; text-align: right;" colspan="4"><strong>Итого:</strong></td>
        <td class="xl72" style="width: 10px; height: 18px; text-align: right;"><strong>$49</strong></td>
        <td class="xl72" style="width: 79px; height: 18px; text-align: right;"><strong>$50</strong></td>
        <td class="xl72" style="width: 77px; height: 18px; text-align: right;"><strong>$51</strong></td>
        <td class="xl72" style="width: 94px; height: 18px; text-align: right;"><strong>$52</strong></td>
        </tr>
        <tr style="height: 20px;">
        <td class="xl77" style="font-size: 15pt; font-family: Bold; vertical-align: middle; text-align: center;  height: 20px; width: 1000px;" colspan="8"><strong>Центр торгово-экономических компетенций</strong></td>
        </tr>
        <tr style="height: 72px;">
        <td class="xl65" style="width: 200px; height: 72px;">Товароведение и экспертиза качества потребительских товаров</td>
        <td class="xl65" style="width: 109px; height: 72px;">38.02.05</td>
        <td class="xl65" style="width: 69px; height: 72px;">очная</td>
        <td class="xl65" style="width: 175px; height: 72px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$53</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$54</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$55</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$56</td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl65" style="width: 200px; height: 18px;">Коммерция</td>
        <td class="xl65" style="width: 109px; height: 18px;">38.02.04</td>
        <td class="xl65" style="width: 69px; height: 18px;">очная</td>
        <td class="xl65" style="width: 175px; height: 18px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$57</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$58</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$59</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$60</td>
        </tr>
        <tr style="height: 72px;">
        <td class="xl65" style="width: 200px; height: 72px;">Товароведение и экспертиза качества потребительских товаров</td>
        <td class="xl65" style="width: 109px; height: 72px;">38.02.05</td>
        <td class="xl65" style="width: 69px; height: 72px;">очно-заочная</td>
        <td class="xl65" style="width: 175px; height: 72px;">11 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$61</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$62</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$63</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$64</td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl66" style="width: 200px; height: 18px; text-align: right;" colspan="4"><strong>Итого:</strong></td>
        <td class="xl72" style="width: 10px; height: 18px; text-align: right;"><strong>$65</strong></td>
        <td class="xl72" style="width: 79px; height: 18px; text-align: right;"><strong>$66</strong></td>
        <td class="xl72" style="width: 77px; height: 18px; text-align: right;"><strong>$67</strong></td>
        <td class="xl72" style="width: 94px; height: 18px; text-align: right;"><strong>$68</strong></td>
        </tr>
        <tr style="height: 20px;">
        <td class="xl77" style="font-size: 15pt; font-family: Bold; vertical-align: middle; text-align: center;  height: 20px; width: 1000px;" colspan="8"><strong>Центр алмазных технологий и геммологии</strong></td>
        </tr>
        <tr style="height: 36px;">
        <td class="xl65" style="width: 200px; height: 36px;">Огранка алмазов в бриллианты</td>
        <td class="xl76" style="width: 109px; height: 36px;">29/01/28</td>
        <td class="xl65" style="width: 69px; height: 36px;">очная</td>
        <td class="xl65" style="width: 175px; height: 36px;">11 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$69</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$70</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$71</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$72</td>
        </tr>
        <tr style="height: 36px;">
        <td class="xl65" style="width: 200px; height: 36px;">Технология обработки алмазов</td>
        <td class="xl76" style="width: 109px; height: 36px;">29/02/08</td>
        <td class="xl65" style="width: 69px; height: 36px;">очная</td>
        <td class="xl65" style="width: 175px; height: 36px;">11 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$73</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$74</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$75</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$76</td>
        </tr>
        <tr style="height: 36px;">
        <td class="xl65" style="width: 200px; height: 36px;">Технология обработки алмазов</td>
        <td class="xl76" style="width: 109px; height: 36px;">29/02/08</td>
        <td class="xl65" style="width: 69px; height: 36px;">очная</td>
        <td class="xl65" style="width: 175px; height: 36px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$77</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$78</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$79</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$80</td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl66" style="width: 200px; height: 18px; text-align: right;" colspan="4"><strong>Итого:&nbsp;&nbsp;&nbsp;</strong></td>
        <td class="xl72" style="width: 10px; height: 18px; text-align: right;"><strong>$81</strong></td>
        <td class="xl72" style="width: 79px; height: 18px; text-align: right;"><strong>$82</strong></td>
        <td class="xl72" style="width: 77px; height: 18px; text-align: right;"><strong>$83</strong></td>
        <td class="xl72" style="width: 94px; height: 18px; text-align: right;"><strong>$84</strong></td>
        </tr>
        <tr style="height: 20px;">
        <td class="xl77" style="font-size: 15pt; font-family: Bold; vertical-align: middle; text-align: center; height: 20px; width: 1000px;" colspan="8"><strong>Центр предпринимательства и развития бизнеса</strong></td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl65" style="width: 200px; height: 18px;">Туризм</td>
        <td class="xl65" style="width: 109px; height: 18px;">43.02.10</td>
        <td class="xl65" style="width: 69px; height: 18px;">очная</td>
        <td class="xl65" style="width: 175px; height: 18px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$85</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$86</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$87</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$88</td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl65" style="width: 200px; height: 18px;">Банковское дело</td>
        <td class="xl65" style="width: 109px; height: 18px;">38.02.07</td>
        <td class="xl65" style="width: 69px; height: 18px;">очная</td>
        <td class="xl65" style="width: 175px; height: 18px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$89</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$90</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$91</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$92</td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl66" style="width: 200px; height: 18px; text-align: right;" colspan="4"><strong>Итого:</strong><strong>&nbsp;</strong><strong>&nbsp;</strong><strong>&nbsp;</strong></td>
        <td class="xl72" style="width: 10px; height: 18px; text-align: right;"><strong>$93</strong></td>
        <td class="xl72" style="width: 79px; height: 18px; text-align: right;"><strong>$94</strong></td>
        <td class="xl72" style="width: 77px; height: 18px; text-align: right;"><strong>$95</strong></td>
        <td class="xl72" style="width: 94px; height: 18px; text-align: right;"><strong>$96</strong></td>
        </tr>
        <tr style="height: 20px;">
        <td class="xl77" style=" font-size: 15pt; font-family: Bold; vertical-align: middle; text-align: center;  height: 20px; width: 1000px;" colspan="8"><strong>Центр аудиовизуальных технологий</strong></td>
        </tr>
        <tr style="height: 36px;">
        <td class="xl65" style="width: 200px; height: 36px;">Аудиовизуальная техника</td>
        <td class="xl76" style="width: 109px; height: 36px;">11/02/05</td>
        <td class="xl65" style="width: 69px; height: 36px;">очная</td>
        <td class="xl65" style="width: 175px; height: 36px;">11 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$97</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$98</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$99</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$100</td>
        </tr>
        <tr style="height: 36px;">
        <td class="xl65" style="width: 200px; height: 36px;">Техника и искусство фотографии</td>
        <td class="xl65" style="width: 109px; height: 36px;">54.02.08</td>
        <td class="xl65" style="width: 69px; height: 36px;">очная</td>
        <td class="xl65" style="width: 175px; height: 36px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$101</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$102</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$103</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$104</td>
        </tr>
        <tr style="height: 54px;">
        <td class="xl69" style="width: 200px; height: 54px;">Музыкальное звукооператорское мастерство</td>
        <td class="xl69" style="width: 109px; height: 54px;">53.02.08</td>
        <td class="xl69" style="width: 69px; height: 54px;">очная</td>
        <td class="xl69" style="width: 175px; height: 54px;">9 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$105</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$106</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$107</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$108</td>
        </tr>
        <tr style="height: 54px;">
        <td class="xl65" style="width: 200px; height: 54px;">Театральная и аудиовизуальная техника (по видам)</td>
        <td class="xl65" style="width: 109px; height: 54px;">55.02.01</td>
        <td class="xl65" style="width: 69px; height: 54px;">очная</td>
        <td class="xl65" style="width: 175px; height: 54px;">11 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$109</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$110</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$111</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$112</td>
        </tr>
        <tr style="height: 36px;">
        <td class="xl65" style="width: 200px; height: 36px;">Анимация (по видам)</td>
        <td class="xl65" style="width: 109px; height: 36px;">55.02.02</td>
        <td class="xl65" style="width: 69px; height: 36px;">очная</td>
        <td class="xl65" style="width: 175px; height: 36px;">11 кл.</td>
        <td class="xl73" style="width: 10px; height: 18px; text-align: right;">$113</td>
        <td class="xl73" style="width: 79px; height: 18px; text-align: right;">$114</td>
        <td class="xl73" style="width: 77px; height: 18px; text-align: right;">$115</td>
        <td class="xl73" style="width: 94px; height: 18px; text-align: right;">$116</td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl66" style="width: 332px; height: 18px; text-align: right;" colspan="4"><strong>Итого:</strong></td>
        <td class="xl72" style="width: 10px; height: 18px; text-align: right;"><strong>$117</strong></td>
        <td class="xl72" style="width: 79px; height: 18px; text-align: right;"><strong>$118</strong></td>
        <td class="xl72" style="width: 77px; height: 18px; text-align: right;"><strong>$119</strong></td>
        <td class="xl72" style="width: 94px; height: 18px; text-align: right;"><strong>$120</strong></td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl66"  style="width: 332px; background-color: light-gray; height: 18px; text-align: right;" colspan="8"><strong></strong></td>
        </tr>
        <tr style="height: 36px;">
        <td class="xl66" style="width: 332px; height: 36px; text-align: right;" colspan="4"><strong>&nbsp;Итого 9кл.:</strong></td>
        <td class="xl72" style="width: 10px; height: 18px; text-align: right;"><strong>$121</strong></td>
        <td class="xl72" style="width: 79px; height: 18px; text-align: right;"><strong>$122</strong></td>
        <td class="xl72" style="width: 77px; height: 18px; text-align: right;"><strong>$123</strong></td>
        <td class="xl72" style="width: 94px; height: 18px; text-align: right;"><strong>$124</strong></td>
        </tr>
        <tr style="height: 37px;">
        <td class="xl67" style="width: 332px; height: 37px; text-align: right;" colspan="4"><strong>&nbsp;Итого 11 кл.:</strong></td>
        <td class="xl72" style="width: 10px; height: 18px; text-align: right;"><strong>$125</strong></td>
        <td class="xl72" style="width: 79px; height: 18px; text-align: right;"><strong>$126</strong></td>
        <td class="xl72" style="width: 77px; height: 18px; text-align: right;"><strong>$127</strong></td>
        <td class="xl72" style="width: 94px; height: 18px; text-align: right;"><strong>$128</strong></td>
        </tr>
        <tr style="height: 18px;">
        <td class="xl67" style="width: 332px; height: 18px; text-align: right;" colspan="4"><strong>&nbsp;Итого:</strong></td>
        <td class="xl72" style="width: 10px; height: 18px; text-align: right;"><strong>$129</strong></td>
        <td class="xl72" style="width: 79px; height: 18px; text-align: right;"><strong>$130</strong></td>
        <td class="xl72" style="width: 77px; height: 18px; text-align: right;"><strong>$131</strong></td>
        <td class="xl72" style="width: 94px; height: 18px; text-align: right;"><strong>$132</strong></td>
        </tr>
        </tbody>
        </table>
"""
}
