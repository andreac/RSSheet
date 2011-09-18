//
//  NSStyle.h
//  excelWriterExample
//
//  Created by andrea cappellotto on 14/09/11.
//  Copyright 2011 Università degli studi di Trento. All rights reserved.
//
typedef enum {
    RSStyleTopAlign = 0,  
    RSStyleCenterAlign = 1,
    RSStyleBottomAlign = 2
}verticalAlign;

typedef enum {
    RSStyleHTopAlign = 0,  
    RSStyleHCenterAlign = 1,
    RSStyleHBottomAlign = 2
}horizontalAlign;

#import <Foundation/Foundation.h>

@interface RSStyle : NSObject
{
    verticalAlign alignmentV;
    horizontalAlign alignmentH;
    UIFont * font;
    UIColor * color;
    float size;
}

@property(nonatomic, assign)verticalAlign alignmentV;
@property(nonatomic, assign)horizontalAlign alignmentH;
@property(nonatomic, retain)UIFont * font;
@property(nonatomic, retain)UIColor * color;
@property(nonatomic, assign)float size;
@end
