//
//  workSheetRow.m
//  excelWriterExample
//
//  Created by andrea cappellotto on 14/09/11.
//  Copyright 2011 Università degli studi di Trento. All rights reserved.
//

#import "RSworkSheetRow.h"

@implementation RSworkSheetRow
@synthesize style, height, cellArray;

- (id)init
{
    self = [super init];
    if (self) {
        // Initialization code here.
    }
    
    return self;
}


- (id)initWithHeight:(NSInteger)heightRow
{
    self = [super init];
    if (self) {
        height = heightRow;
        style = [[RSStyle alloc] init];
        style.font = [UIFont systemFontOfSize:14];
        style.size = 14;
        style.color = [UIColor blackColor];
        style.alignmentH = RSStyleMiddleAlign;
        style.alignmentV = RSStyleCenterAlign;
        
        cellArray = [[NSMutableArray alloc] init];
    }
    
    return self;
}
- (id)initiWithHeight:(NSInteger)heightRow andStyle:(RSStyle *)styleRow
{
    self = [super init];
    if (self) {
        style = [styleRow retain];
        height = heightRow;
        cellArray = [[NSMutableArray alloc] init];
    }
    
    return self; 
}

- (void)addCellString:(NSString *)contentRow
{
    RSStyle * styleDefault = [[RSStyle alloc] init];
    styleDefault.font = [UIFont systemFontOfSize:14];
    styleDefault.size = 14;
    styleDefault.color = [UIColor blackColor];
    styleDefault.alignmentH = RSStyleMiddleAlign;
    styleDefault.alignmentV = RSStyleCenterAlign;
    
    RSCell * newCell = [[RSCell alloc] init];
    newCell.content = contentRow;
    newCell.type = cellTypeString;
    newCell.style = styleDefault;
    
    [styleDefault release];
    
    [cellArray addObject:newCell];
    [newCell release];
    
}
- (void)addCellString:(NSString *)contentRow withStyle:(RSStyle *)styleRow
{
    RSCell * newCell = [[RSCell alloc] init];
    newCell.content = contentRow;
    newCell.type = cellTypeString;
    newCell.style = styleRow;
    
    [cellArray addObject:newCell];
    [newCell release];
}
- (void)addCellNumber:(float)contentRow
{
    RSStyle * styleDefault = [[RSStyle alloc] init];
    styleDefault.font = [UIFont systemFontOfSize:14];
    styleDefault.size = 14;
    styleDefault.color = [UIColor blackColor];
    styleDefault.alignmentH = RSStyleMiddleAlign;
    styleDefault.alignmentV = RSStyleCenterAlign;
    
    RSCell * newCell = [[RSCell alloc] init];
    newCell.content = [NSString stringWithFormat:@"%.2f", contentRow];
    newCell.type = cellTypeNumber;
    newCell.style = styleDefault;
    [styleDefault release];
    
    [cellArray addObject:newCell];
    [newCell release];
}
- (void)addCellNumber:(float)contentRow withStyle:(RSStyle *)styleRow
{
    RSCell * newCell = [[RSCell alloc] init];
    newCell.content = [NSString stringWithFormat:@"%.2f", contentRow];
    newCell.type = cellTypeNumber;
    newCell.style = styleRow;
    
    [cellArray addObject:newCell];
    [newCell release];
}
- (void)addCellData:(NSDate *)contentRow
{
    NSDateFormatter *format = [[NSDateFormatter alloc] init];
    [format setDateFormat:@"yyyy-MM-ddTHH:mm:ss"];
    NSString *dateString = [format stringFromDate:contentRow];
    [format release];
    
    RSStyle * styleDefault = [[RSStyle alloc] init];
    styleDefault.font = [UIFont systemFontOfSize:14];
    styleDefault.size = 14;
    styleDefault.color = [UIColor blackColor];
    styleDefault.alignmentH = RSStyleMiddleAlign;
    styleDefault.alignmentV = RSStyleCenterAlign;
    
    RSCell * newCell = [[RSCell alloc] init];
    newCell.content = dateString;
    newCell.type = cellTypeDate;
    newCell.style = styleDefault;
    [styleDefault release];
    
    [cellArray addObject:newCell];
    [newCell release];
    
    
}
- (void)addCellData:(NSDate *)contentRow withStyle:(RSStyle *)styleRow
{
    NSDateFormatter *format = [[NSDateFormatter alloc] init];
    [format setDateFormat:@"yyyy-MM-ddTHH:mm:ss"];
    NSString *dateString = [format stringFromDate:contentRow];
    [format release];
    
    RSCell * newCell = [[RSCell alloc] init];
    newCell.content = dateString;
    newCell.type = cellTypeDate;
    newCell.style = styleRow;
    
    [cellArray addObject:newCell];
    [newCell release];
}

-(void)dealloc
{
    [super dealloc];
    [cellArray release];
    [style release];
}


@end
