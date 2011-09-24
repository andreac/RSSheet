//
//  workBook.h
//  excelWriterExample
//
//  Created by andrea cappellotto on 14/09/11.
//  Copyright 2011 Università degli studi di Trento. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "UIColor-Expanded.h"

#import "RSworkSheet.h"
#import "RSworkSheetRow.h"
#import "RSCell.h"
#import "RSStyle.h"
@interface RSworkBook : NSObject
{
    NSString * author;
    float version;
    NSDate * date;
    RSStyle * defaultStyle;
    
    NSMutableArray * arrayWorkSheet;
    
}

@property(nonatomic, retain)NSString * author;
@property(nonatomic, retain)NSDate * date;
@property(nonatomic, assign)float version;
@property(nonatomic, retain)RSStyle * defaultStyle;

- (void)addWorkSheet:(RSworkSheet *)sheet;
- (BOOL)writeWithName:(NSString*)name toPath:(NSString*)path;
@end
