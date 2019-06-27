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
                
                if row.cells.count == 7 {
                    continue
                }
                
                if row.cells.count == 9 {
                    var nilCount = 0
                    for cell in row.cells {
                        if cell.value == nil {
                            nilCount += 1
                           
                        }
                    }
                    
                    if nilCount >= 8 {
                        continue
                    }
                }
                
                if row.cells.count == 3 {
                    var nilCount = 0
                    for cell in row.cells {
                        if cell.value == nil {
                            nilCount += 1
                        }
                    }
                    
                  
                }
                print("Count: \(row.cells.count)")
                for cell in row.cells {

                    if collumnCount < 5 {
                     collumnCount += 1
                        continue
                    }
                    
                    if row.cells.count >= 10 {
                        if cell == row.cells[9]{
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
    <tr>
        <td><strong>Специальность</strong></td>
        <td><strong>Название группы</strong></td>
        <td><strong>Код специальности</strong></td>
        <td><strong>Форма обучения</strong></td>
        <td><strong>База приема</strong></td>
        <td><strong>План набора (бюджет)</strong></td>
        <td><strong>План</strong><br /><strong>набора (внебюджет)</strong></td>
        <td><strong>Подано заявлений (бюджет)</strong></td>
        <td><strong>Подано заявлений (внебюджет)</strong></td>
    </tr>
    <tr>
        <td style="font-family: Bold; vertical-align: middle; text-align: center; " colspan="9"><strong>Центр медицинской техники и оптики</strong></td>
    </tr>
    <tr>
        <td>Медицинская оптика</td>
        <td>МО-11</td>
        <td>31.02.04</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$1</td>
        <td style="text-align: right;">$2</td>
        <td style="text-align: right;">$3</td>
        <td style="text-align: right;">$4</td>
    </tr>
    <tr>
        <td>Медицинская оптика</td>
        <td>МО-12</td>
        <td>31.02.04</td>
        <td>очная</td>
        <td>11 кл.</td>
        <td style="text-align: right;">$5</td>
        <td style="text-align: right;">$6</td>
        <td style="text-align: right;">$7</td>
        <td style="text-align: right;">$8</td>
    </tr>
    <tr>
        <td>Монтаж, техническое обслуживание и ремонт медицинской техники</td>
        <td>МТ-11</td>
        <td>12/02/07</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$9</td>
        <td style="text-align: right;">$10</td>
        <td style="text-align: right;">$11</td>
        <td style="text-align: right;">$12</td>
    </tr>
    <tr>
        <td>Аддитивные технологии</td>
        <td>АТ-11</td>
        <td>15/02/09</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$13</td>
        <td style="text-align: right;">$14</td>
        <td style="text-align: right;">$15</td>
        <td style="text-align: right;">$16</td>
    </tr>
    <tr>
        <td>Медицинская оптика (УП)</td>
        <td>ВМО-27</td>
        <td>31.02.04</td>
        <td>заочная</td>
        <td>11 кл.</td>
        <td style="text-align: right;">$17</td>
        <td style="text-align: right;">$18</td>
        <td style="text-align: right;">$19</td>
        <td style="text-align: right;">$20</td>
    </tr>
    <tr>
        <td>Медицинская оптика (БП)</td>
        <td>ВМО-11, ВМО-12, ВМО-13</td>
        <td>31.02.04</td>
        <td>заочная</td>
        <td>11 кл.&nbsp;</td>
        <td style="text-align: right;">$21</td>
        <td style="text-align: right;">$22</td>
        <td style="text-align: right;">$23</td>
        <td style="text-align: right;">$24</td>
    </tr>
    <tr>
        <td style="text-align: right;" colspan="5"><strong>Итого:</strong></td>
        <td style="text-align: right;"><strong>$25</strong></td>
        <td style="text-align: right;"><strong>$26</strong></td>
        <td style="text-align: right;"><strong>$27</strong></td>
        <td style="text-align: right;"><strong>$28</strong></td>
    </tr>
    <tr>
        <td style="font-family: Bold; vertical-align: middle; text-align: center; " colspan="9"><strong>Центр информационно-коммуникационных технологий</strong></td>
    </tr>
    <tr>
        <td>Информационные системы и программирование</td>
        <td>ИСИП-11</td>
        <td>09/02/07</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$29</td>
        <td style="text-align: right;">$30</td>
        <td style="text-align: right;">$31</td>
        <td style="text-align: right;">$32</td>
    </tr>
    <tr>
        <td>Информационные системы и программирование</td>
        <td>ИСИП-14</td>
        <td>09/02/07</td>
        <td>очная</td>
        <td>11 кл.</td>
        <td style="text-align: right;">$33</td>
        <td style="text-align: right;">$34</td>
        <td style="text-align: right;">$35</td>
        <td style="text-align: right;">$36</td>
    </tr>
    <tr>
        <td>Обеспечение информационной безопасности автоматизированных систем</td>
        <td>ИБ-12</td>
        <td>10/02/05</td>
        <td>очная</td>
        <td><p>11 кл.</p></td>
        <td style="text-align: right;">$37</td>
        <td style="text-align: right;">$38</td>
        <td style="text-align: right;">$39</td>
        <td style="text-align: right;">$40</td>
    </tr>
    <tr>
        <td>Сетевое и системное администрирование</td>
        <td>С-11</td>
        <td>09/02/06</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$41</td>
        <td style="text-align: right;">$42</td>
        <td style="text-align: right;">$43</td>
        <td style="text-align: right;">$44</td>
    </tr>
    <tr>
        <td>Компьютерные системы и комплексы</td>
        <td>КСИК-12</td>
        <td>09/02/01</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$45</td>
        <td style="text-align: right;">$46</td>
        <td style="text-align: right;">$47</td>
        <td style="text-align: right;">$48</td>
    </tr>
    <tr>
        <td style="text-align: right;" colspan="5"><strong>Итого:</strong></td>
        <td style="text-align: right;"><strong>$49</strong></td>
        <td style="text-align: right;"><strong>$50</strong></td>
        <td style="text-align: right;"><strong>$51</strong></td>
        <td style="text-align: right;"><strong>$52</strong></td>
    </tr>
    <tr>
        <td style="font-family: Bold; vertical-align: middle; text-align: center;   " colspan="89"><strong>Центр торгово-экономических компетенций</strong></td>
    </tr>
    <tr>
        <td>Товароведение и экспертиза качества потребительских товаров</td>
        <td>Т-15</td>
        <td>38.02.05</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$53</td>
        <td style="text-align: right;">$54</td>
        <td style="text-align: right;">$55</td>
        <td style="text-align: right;">$56</td>
    </tr>
    <tr>
        <td>Коммерция</td>
        <td>КМ-15</td>
        <td>38.02.04</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$57</td>
        <td style="text-align: right;">$58</td>
        <td style="text-align: right;">$59</td>
        <td style="text-align: right;">$60</td>
    </tr>
    <tr>
        <td>Товароведение и экспертиза качества потребительских товаров</td>
        <td>ТЭОЗ-13</td>
        <td>38.02.05</td>
        <td>очно-заочная</td>
        <td>11 кл.</td>
        <td style="text-align: right;">$61</td>
        <td style="text-align: right;">$62</td>
        <td style="text-align: right;">$63</td>
        <td style="text-align: right;">$64</td>
    </tr>
    <tr>
        <td style="text-align: right;" colspan="5"><strong>Итого:</strong></td>
        <td style="text-align: right;"><strong>$65</strong></td>
        <td style="text-align: right;"><strong>$66</strong></td>
        <td style="text-align: right;"><strong>$67</strong></td>
        <td style="text-align: right;"><strong>$68</strong></td>
    </tr>
    <tr>
        <td style="font-family: Bold; vertical-align: middle; text-align: center;   " colspan="9"><strong>Центр алмазных технологий и геммологии</strong></td>
    </tr>
    <tr>
        <td>Огранка алмазов в бриллианты</td>
        <td>ОА-11, ОА-12</td>
        <td>29/01/28</td>
        <td>очная</td>
        <td>11 кл.</td>
        <td style="text-align: right;">$69</td>
        <td style="text-align: right;">$70</td>
        <td style="text-align: right;">$71</td>
        <td style="text-align: right;">$72</td>
    </tr>
    <tr>
        <td>Технология обработки алмазов</td>
        <td>ТОА-18</td>
        <td>29/02/08</td>
        <td>очная</td>
        <td>11 кл.</td>
        <td style="text-align: right;">$73</td>
        <td style="text-align: right;">$74</td>
        <td style="text-align: right;">$75</td>
        <td style="text-align: right;">$76</td>
    </tr>
    <tr>
        <td>Технология обработки алмазов</td>
        <td>ТОА-14, ТОА-15</td>
        <td>29/02/08</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$77</td>
        <td style="text-align: right;">$78</td>
        <td style="text-align: right;">$79</td>
        <td style="text-align: right;">$80</td>
    </tr>
    <tr>
        <td style="text-align: right;" colspan="5"><strong>Итого:</strong></td>
        <td style="text-align: right;"><strong>$81</strong></td>
        <td style="text-align: right;"><strong>$82</strong></td>
        <td style="text-align: right;"><strong>$83</strong></td>
        <td style="text-align: right;"><strong>$84</strong></td>
    </tr>
    <tr>
        <td style="font-family: Bold; vertical-align: middle; text-align: center;  " colspan="9"><strong>Центр предпринимательства и развития бизнеса</strong></td>
    </tr>
    <tr>
        <td>Туризм</td>
        <td>ТР-16,ТР-17, ТР-18</td>
        <td>43.02.10</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$85</td>
        <td style="text-align: right;">$86</td>
        <td style="text-align: right;">$87</td>
        <td style="text-align: right;">$88</td>
    </tr>
    <tr>
        <td>Банковское дело</td>
        <td>БД-14</td>
        <td>38.02.07</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$89</td>
        <td style="text-align: right;">$90</td>
        <td style="text-align: right;">$91</td>
        <td style="text-align: right;">$92</td>
    </tr>
    <tr>
        <td style="text-align: right;" colspan="5"><strong>Итого:</strong><strong>&nbsp;</strong><strong>&nbsp;</strong><strong>&nbsp;</strong></td>
        <td style="text-align: right;"><strong>$93</strong></td>
        <td style="text-align: right;"><strong>$94</strong></td>
        <td style="text-align: right;"><strong>$95</strong></td>
        <td style="text-align: right;"><strong>$96</strong></td>
    </tr>
    <tr>
        <td style=" font-family: Bold; vertical-align: middle; text-align: center;   " colspan="9"><strong>Центр аудиовизуальных технологий</strong></td>
    </tr>
    <tr>
        <td>Аудиовизуальная техника</td>
        <td>АВТ-11</td>
        <td>11/02/05</td>
        <td>очная</td>
        <td>11 кл.</td>
        <td style="text-align: right;">$97</td>
        <td style="text-align: right;">$98</td>
        <td style="text-align: right;">$99</td>
        <td style="text-align: right;">$100</td>
    </tr>
    <tr>
        <td>Техника и искусство фотографии</td>
        <td>ФВТ-11</td>
        <td>54.02.08</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$101</td>
        <td style="text-align: right;">$102</td>
        <td style="text-align: right;">$103</td>
        <td style="text-align: right;">$104</td>
    </tr>
    <tr>
        <td>Музыкальное звукооператорское мастерство</td>
        <td>МЗМ-11</td>
        <td>53.02.08</td>
        <td>очная</td>
        <td>9 кл.</td>
        <td style="text-align: right;">$105</td>
        <td style="text-align: right;">$106</td>
        <td style="text-align: right;">$107</td>
        <td style="text-align: right;">$108</td>
    </tr>
    <tr>
        <td>Анимация (по видам)</td>
        <td>А-12</td>
        <td>55.02.02</td>
        <td>очная</td>
        <td>11 кл.</td>
        <td style="text-align: right;">$109</td>
        <td style="text-align: right;">$110</td>
        <td style="text-align: right;">$111</td>
        <td style="text-align: right;">$112</td>
    </tr>
    <tr>
        <td>Театральная и аудиовизуальная техника (по видам)</td>
        <td>ТАТ-12</td>
        <td>55.02.01</td>
        <td>очная</td>
        <td>11 кл.</td>
        <td style="text-align: right;">$113</td>
        <td style="text-align: right;">$114</td>
        <td style="text-align: right;">$115</td>
        <td style="text-align: right;">$116</td>
    </tr>
    <tr>
        <td style="text-align: right;" colspan="5"><strong>Итого:</strong></td>
        <td style="text-align: right;"><strong>$117</strong></td>
        <td style="text-align: right;"><strong>$118</strong></td>
        <td style="text-align: right;"><strong>$119</strong></td>
        <td style="text-align: right;"><strong>$120</strong></td>
    </tr>
    <tr>
        <td  style=" background-color: lightgray;text-align: right;" colspan="9"><strong></strong></td>
    </tr>
    <tr>
        <td style="text-align: right;" colspan="5"><strong>&nbsp;Итого 9кл.:</strong></td>
        <td style="text-align: right;"><strong>$121</strong></td>
        <td style="text-align: right;"><strong>$122</strong></td>
        <td style="text-align: right;"><strong>$123</strong></td>
        <td style="text-align: right;"><strong>$124</strong></td>
    </tr>
    <tr>
        <td style="text-align: right;" colspan="5"><strong>&nbsp;Итого 11 кл.:</strong></td>
        <td style="text-align: right;"><strong>$125</strong></td>
        <td style="text-align: right;"><strong>$126</strong></td>
        <td style="text-align: right;"><strong>$127</strong></td>
        <td style="text-align: right;"><strong>$128</strong></td>
    </tr>
    <tr>
        <td style="text-align: right;" colspan="5"><strong>&nbsp;Итого:</strong></td>
        <td style="text-align: right;"><strong>$129</strong></td>
        <td style="text-align: right;"><strong>$130</strong></td>
        <td style="text-align: right;"><strong>$131</strong></td>
        <td style="text-align: right;"><strong>$132</strong></td>
    </tr>
    </tbody>
</table>
"""
}
