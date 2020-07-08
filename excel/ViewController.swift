//
//  ViewController.swift
//  excel
//
//  Created by 飯沼圭哉 on 2020/07/08.
//  Copyright © 2020 keisuke.iinuma. All rights reserved.
//

import UIKit
import CoreXLSX

class ViewController: UIViewController {
    
    override func viewDidLoad() {
        super.viewDidLoad()
        // Do any additional setup after loading the view.
        
        let filepath = "./categories.xlsx"
        guard let file = XLSXFile(filepath: filepath) else {
            fatalError("XLSX file at \(filepath) is corrupted or does not exist")
        }
        
        for wbk in try file.parseWorkbooks() {
            for (name, path) in try file.parseWorksheetPathsAndNames(workbook: wbk) {
                if let worksheetName = name {
                    print("This worksheet has a name: \(worksheetName)")
                }
                
                let worksheet = try file.parseWorksheet(at: path)
                for row in worksheet.data?.rows ?? [] {
                    for c in row.cells {
                        print(c)
                    }
                }
            }
        }
        
    }
    
}

