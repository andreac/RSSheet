//
//  workSheetRow.h
//  excelWriterExample
//
//  Created by andrea cappellotto on 14/09/11.
//  Copyright 2011 Università degli studi di Trento. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "RSStyle.h"
#import "RSCell.h"

@interface RSworkSheetRow : NSObject
{
    NSMutableArray * cellArray;
    float height;
    RSStyle * style;
}
@property(nonatomic, readonly)NSMutableArray* cellArray;
@property(nonatomic, assign)float height;
@property(nonatomic, retain)RSStyle * style;


- (id)initWithHeight:(NSInteger )height;
- (id)initiWithHeight:(NSInteger )height andStyle:(RSStyle *)style;

- (void)addCellString:(NSString *)content;
- (void)addCellString:(NSString *)content withStyle:(RSStyle *)style;
- (void)addCellNumber:(float )content;
- (void)addCellNumber:(float )content withStyle:(RSStyle *)style;
- (void)addCellData:(NSDate *)content;
- (void)addCellData:(NSDate *)content withStyle:(RSStyle *)style;
@end
