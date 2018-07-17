import * as React from 'react';

interface IErrorBoundaryState
{
    error: string;
    errorInfo: any;
}

export class ErrorBoundary extends React.Component<{}, IErrorBoundaryState> {
    constructor(props) 
    {
      super(props);
      this.state = { error: null, errorInfo: null };
    }
    
    componentDidCatch(error, errorInfo) 
    {
      this.setState(
        {
            error: error,
            errorInfo: errorInfo
        });
    }
    
    render() 
    {
        // If an unhandled error occurred then display it to the user
        if (this.state.errorInfo) 
        {
            return (
                <div>
                    <h2>An error occurred.</h2>
                    <details style={{ whiteSpace: 'pre-wrap' }}>
                    {this.state.error && this.state.error.toString()}
                    <br />
                    {this.state.errorInfo.componentStack}
                    </details>
                </div>
            );
        }
      
      // Happy path
      return this.props.children;
    }  
  }