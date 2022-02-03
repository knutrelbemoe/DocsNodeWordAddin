import * as React from 'react';
import {
  DocumentCard,  
  DocumentCardTitle,
  DocumentCardPreview,
  IDocumentCardPreviewProps 
} from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import {IDocsNodeRenderImageListProp} from './IDocsNodeRenderImageListProps';
import styles from './DocsNodeAdmin.module.scss';

export default class DocsNodeRenderImageList  extends React.Component<IDocsNodeRenderImageListProp>{
    
    constructor(props: IDocsNodeRenderImageListProp) {
        super(props); 
    }
    public render():React.ReactElement<IDocsNodeRenderImageListProp> {
        const previewProps: IDocumentCardPreviewProps = {
            previewImages: [
              {
                name: 'Revenue stream proposal fiscal year 2016 version02.pptx',                
                previewImageSrc: this.props.imageItems.ImageUrl, 
                iconSrc: '',               
                imageFit: ImageFit.contain,
                width: 210,
                height: 100
              }
            ]
          };
        return (
            <div>
            <DocumentCard className={styles.DocumentView}>
              <div>
                <DocumentCardPreview {...previewProps} />
              </div>              
              <DocumentCardTitle
                title={this.props.imageItems.Title}
                shouldTruncate={true}
              />
            </DocumentCard>
            </div>
        );
    }
}