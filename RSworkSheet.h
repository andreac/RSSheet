//
//  workSheet.h
//  excelWriterExample
//
//  Created by andrea cappellotto on 14/09/11.
//  Copyright 2011 Università degli studi di Trento. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "RSworkSheetRow.h"

@interface RSworkSheet : NSObject
{
    float columnWidth;
    float rowHeight;
    NSString * name;
    NSMutableArray * arrayWorkSheetRow;
}

@property(nonatomic, assign)float columnWidth;
@property(nonatomic, assign)float rowHeight;
@property(nonatomic, retain)NSString * name;
@property(nonatomic, readonly)NSMutableArray * arrayWorkSheetRow;

- (void)addWorkSheetRow:(RSworkSheetRow*)row;
- (id)initWithName:(NSString*)nameSheet;
@end
