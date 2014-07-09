//
//  workBook.m
//  excelWriterExample
//
//  Created by andrea cappellotto on 14/09/11.
//  Copyright 2011 Università degli studi di Trento. All rights reserved.
//

#import "RSworkBook.h"

@interface RSworkBook()
- (NSString*)formatTypeToStringVertical:(verticalAlign)formatType;
- (NSString*)formatTypeToStringHorizontal:(horizontalAlign)formatType;
- (NSString*)formatTypeToCellType:(cellType)formatType;
- (int)countMaxCell:(NSMutableArray*)sheet;
-(BOOL)isEqualStyle:(RSStyle*)style1 toStyle:(RSStyle*)style2;
@end


@implementation RSworkBook

@synthesize version, author, date, defaultStyle;
- (id)init
{
    self = [super init];
    if (self) 
    {
        defaultStyle = [[RSStyle alloc] init];
        defaultStyle.font = [UIFont systemFontOfSize:14];
        defaultStyle.size = 14;
        defaultStyle.color = [UIColor blackColor];
        defaultStyle.alignmentH = RSStyleMiddleAlign;
        defaultStyle.alignmentV = RSStyleCenterAlign;
        
        version = 1.0;
        author = [[NSString alloc] initWithFormat:@"system"];
        date = [NSDate date];
        
        arrayWorkSheet = [[NSMutableArray alloc] init];
        
    }
    
    return self;
}


- (void)addWorkSheet:(RSworkSheet *)sheet
{
    [arrayWorkSheet addObject:sheet];
}


- (BOOL)writeWithName:(NSString*)name toPath:(NSString*)path
{
    
    NSDateFormatter *format = [[NSDateFormatter alloc] init];
    [format setDateFormat:@"yyyy-MM-ddTHH:mm:ss"];
    NSString *dateString = [format stringFromDate:self.date];
   
    
    NSString * head = [[NSString alloc] initWithFormat:@"<?xml version=\"1.0\"encoding=\"UTF-8\"?>\n<?mso-application progid=\"Excel.Sheet\"?>\n<Workbook xmlns:c=\"urn:schemas-microsoft-com:office:component:spreadsheet\"\n xmlns:html=\"http://www.w3.org/TR/REC-html40\"\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\n xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n xmlns:x2=\"http://schemas.microsoft.com/office/excel/2003/xml\"\n xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\"><Author>%@</Author>\n<LastAuthor>%@</LastAuthor>\n<Created>%@</Created>\n<Version>%.2f</Version>\n</DocumentProperties>\n", self.author, self.author, dateString, self.version];
    

    NSString * officeSetting = [[NSString alloc] initWithFormat:@"<OfficeDocumentSettings xmlns=\"urn:schemas-microsoft-com:office:office\">\n</OfficeDocumentSettings>\n<ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\">\n<WindowHeight>20000</WindowHeight>\n<WindowWidth>20000</WindowWidth>\n<WindowTopX>0</WindowTopX>\n<WindowTopY>0</WindowTopY>\n<ProtectStructure>False</ProtectStructure>\n<ProtectWindows>False</ProtectWindows>\n</ExcelWorkbook>"];
    
    NSMutableString * styleDefault = [[NSMutableString alloc] initWithFormat:@"<Styles>"];
    
    [styleDefault appendFormat:@"<Style ss:ID=\"Default\" ss:Name=\"Normal\">\n<Alignment ss:Vertical=\"%@\" ss:Horizontal=\"%@\"/>\n<Borders/>\n<Font ss:FontName=\"%@\" ss:Size=\"%.2f\" ss:Color=\"#%@\"/>\n<Interior/>\n<NumberFormat/>\n<Protection/>\n</Style>\n",
     [self formatTypeToStringVertical:self.defaultStyle.alignmentV], 
     [self formatTypeToStringHorizontal:self.defaultStyle.alignmentH], 
     defaultStyle.font.fontName,
     self.defaultStyle.size, 
     [self.defaultStyle.color hexStringFromColor]];
    
    
    BOOL dateType = FALSE;
    int i = 0;
    for (RSworkSheet * sheet in arrayWorkSheet) 
    {
        
        for (RSworkSheetRow * sheetRow in sheet.arrayWorkSheetRow) 
        {
            dateType = FALSE;
            if (![self isEqualStyle:sheetRow.style toStyle:defaultStyle]) 
            {
                [styleDefault appendFormat:@"<Style ss:ID=\"s%i\">\n<Alignment ss:Vertical=\"%@\" ss:Horizontal=\"%@\"/>\n<Borders/>\n<Font ss:FontName=\"%@\" ss:Size=\"%.2f\" ss:Color=\"#%@\"/>\n<Interior/>\n<NumberFormat/>\n<Protection/>\n</Style>\n",
                 60+i,
                 [self formatTypeToStringVertical: sheetRow.style.alignmentV], 
                 [self formatTypeToStringHorizontal:sheetRow.style.alignmentH], 
                 sheetRow.style.font.fontName,
                 sheetRow.style.size, 
                 [sheetRow.style.color hexStringFromColor]];
                
                i++;
                
            }
            for (RSCell * cell in sheetRow.cellArray) 
            {
                if (![self isEqualStyle:cell.style toStyle:defaultStyle]) 
                {
                    [styleDefault appendFormat:@"<Style ss:ID=\"s%i\">\n<Alignment ss:Vertical=\"%@\" ss:Horizontal=\"%@\"/>\n<Borders/>\n<Font ss:FontName=\"%@\" ss:Size=\"%.2f\" ss:Color=\"#%@\"/>\n<Interior/>\n<NumberFormat/>\n<Protection/>\n</Style>\n",
                     60+i,
                     [self formatTypeToStringVertical: sheetRow.style.alignmentV], 
                     [self formatTypeToStringHorizontal:sheetRow.style.alignmentH], 
                     sheetRow.style.font.fontName,
                     sheetRow.style.size, 
                     [sheetRow.style.color hexStringFromColor]];
                    i++;
                }
                if ((cell.type == cellTypeDate)&&(![self isEqualStyle:sheetRow.style toStyle:defaultStyle]))
                {
                    [styleDefault appendFormat:@"<Style ss:ID=\"s%i\">\n<Alignment ss:Vertical=\"%@\" ss:Horizontal=\"%@\"/>\n<Borders/>\n<Font ss:FontName=\"%@\" ss:Size=\"%.2f\" ss:Color=\"#%@\"/>\n<Interior/>\n<NumberFormat ss:Format=\"Short Date\" />\n<Protection/>\n</Style>\n",
                    
                    60+i,
                    [self formatTypeToStringVertical: sheetRow.style.alignmentV], 
                    [self formatTypeToStringHorizontal:sheetRow.style.alignmentH], 
                    sheetRow.style.font.fontName,
                    sheetRow.style.size, 
                    [sheetRow.style.color hexStringFromColor]];
                    dateType = TRUE;
                    
                    i++;
                }
                else if(cell.type == cellTypeDate)
                {
                    [styleDefault  appendFormat:@"<Style ss:ID=\"s%i\">\n<NumberFormat ss:Format=\"Short Date\"/>\n</Style>\n", 60+i];
                    i++; 
                }
            }
        }
    }
    [styleDefault appendFormat:@"</Styles>\n"];
    
    
    NSMutableString * sheet = [[NSMutableString alloc] init];
    
    
    int j = 0;
    for (RSworkSheet * sheetwork in arrayWorkSheet) 
    {
        [sheet appendFormat:@"<Worksheet ss:Name=\"%@\">\n", sheetwork.name];
        
        [sheet appendFormat:@"<Table ss:ExpandedColumnCount=\"%i\" ss:ExpandedRowCount=\"%i\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultColumnWidth=\"%.2f\" ss:DefaultRowHeight=\"%.2f\">\n",
         [self countMaxCell:sheetwork.arrayWorkSheetRow],
         [sheetwork.arrayWorkSheetRow count],
         
         sheetwork.columnWidth, 
         sheetwork.rowHeight];
        
        [sheet appendFormat:@"<Column ss:Width=\"80\"/>\n"];
        
        for (RSworkSheetRow * sheetRow in sheetwork.arrayWorkSheetRow) 
        {
            if (![self isEqualStyle:sheetRow.style toStyle:defaultStyle]) 
            {
                [sheet appendFormat:@"<Row ss:AutoFitHeight=\"0\" ss:Height=\"%.2f\" ss:StyleID=\"s%i\">\n", sheetRow.height, 60+j];
                j++;
            }
            else
            {
                [sheet appendFormat:@"<Row ss:AutoFitHeight=\"0\" ss:Height=\"%.2f\" >\n", sheetRow.height];
            }
            
            
            for (RSCell * cell in sheetRow.cellArray) 
            {
                if (![self isEqualStyle:cell.style toStyle:defaultStyle]) 
                {
                    [sheet appendFormat:@"<Cell ss:StyleID=\"s%i\">\n", 60+j];
                    j++;
                }
                else if (cell.type == cellTypeDate) 
                {
                    [sheet appendFormat:@"<Cell ss:StyleID=\"s%i\">\n", 60+j];
                    j++;
                }
                else
                {
                    [sheet appendFormat:@"<Cell>\n"];
                }
                
                [sheet appendFormat:@"<Data ss:Type=\"%@\">%@</Data>\n", [self formatTypeToCellType:cell.type], cell.content];
                
                [sheet appendFormat:@"</Cell>\n"];
            }
            
            [sheet appendFormat:@"</Row>\n"];
        }
        [sheet appendFormat:@"</Table>\n<WorksheetOptions/>\n</Worksheet>\n"];
    }
    [sheet appendFormat:@"</Workbook>\n"];
    
    
    NSMutableString * finalText = [[NSMutableString alloc] init];
    
    [finalText appendFormat:@"%@",head];
    [finalText appendFormat:@"%@",officeSetting];
    [finalText appendFormat:@"%@",styleDefault];
    [finalText appendFormat:@"%@",sheet];
    
    
    NSLog(@"%@", finalText);
    
    NSString *pathFinal = [path stringByAppendingPathComponent:[NSString stringWithFormat:@"%@.xls",name]];
    
    NSError *err;
    if([finalText writeToFile:pathFinal atomically:YES encoding:NSUTF8StringEncoding error:&err])
    {
        NSLog(@"salvataggio corretto");
        NSLog(@"path:%@", pathFinal);
        
        
        return YES;
    }
    else
    {
        NSLog(@"error:%@", err);
        
        
        return NO;
    }
}
   

-(BOOL)isEqualStyle:(RSStyle*)style1 toStyle:(RSStyle*)style2
{
    if (![style1.font isEqual:style2.font])
        return FALSE;
    if (![style1.color isEqual:style2.color])
        return FALSE;
    if (style1.size != style2.size)
        return FALSE;
    if (style1.alignmentH != style2.alignmentH)
        return FALSE;
    if (style1.alignmentV != style2.alignmentV)
        return FALSE;
    
    return TRUE;
}


- (int)countMaxCell:(NSMutableArray*)sheet
{
    int max=0;
    for (RSworkSheetRow * row in sheet) 
    {
        if ([row.cellArray count] > max) 
        {
            max = [row.cellArray count];
        }
    }
    return max;
}
     
- (NSString*)formatTypeToStringVertical:(verticalAlign)formatType { 
    NSString *result = nil; 
    switch(formatType) { 
        case RSStyleTopAlign: 
            result = @"Top"; 
            break; 
        case RSStyleCenterAlign: 
            result = @"Center"; 
            break; 
        case RSStyleBottomAlign: 
            result = @"Bottom"; 
            break;
        default: 
            [NSException raise:NSGenericException format:@"Unexpected FormatType."]; 
    } 
    return result; 
}

- (NSString*)formatTypeToStringHorizontal:(horizontalAlign)formatType { 
    NSString *result = nil; 
    switch(formatType) { 
        case RSStyleLeftAlign: 
            result = @"Left"; 
            break; 
        case RSStyleMiddleAlign: 
            result = @"Center"; 
            break; 
        case RSStyleRightAlign: 
            result = @"Right"; 
            break;
        default: 
            [NSException raise:NSGenericException format:@"Unexpected FormatType."]; 
    } 
    return result; 
}

- (NSString*)formatTypeToCellType:(cellType)formatType
{
    NSString *result = nil; 
    switch(formatType) { 
        case cellTypeNumber: 
            result = @"Number"; 
            break; 
        case cellTypeString: 
            result = @"String"; 
            break; 
        case cellTypeDate: 
            result = @"DateTime"; 
            break;
        default: 
            [NSException raise:NSGenericException format:@"Unexpected FormatType."]; 
    } 
    return result; 
}
@end
