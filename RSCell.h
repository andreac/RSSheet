//
//  NSCell.h
//  excelWriterExample
//
//  Created by andrea cappellotto on 15/09/11.
//  Copyright 2011 Università degli studi di Trento. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "RSStyle.h"

typedef enum
{
    cellTypeNumber=0,
    cellTypeString=1,
    cellTypeDate=2
    
}cellType;


@interface RSCell : NSObject
{
    RSStyle * style;
    NSString * content;
    cellType type;
}

@property(nonatomic, retain)RSStyle * style;
@property(nonatomic, retain)NSString * content;
@property(nonatomic, assign)cellType type;

@end
