import * as $ from 'jquery';

export class ApiService
{
    private token: string;
    private endpoint: string;
        
    constructor(token: string)
    {
        this.token = token;
        this.endpoint = `${window.location.origin}/api`;
    }

    public callApi(controller, action, method, args)
    {
        return new Promise((resolve, reject) =>
        {
            let endpoint = `${this.endpoint}/${controller}/${action}`;
            $.ajax({
                url: endpoint,
                headers: { "Authorization": "Bearer " + this.token },
                type: method,
                data: args
            })
            .done(response => {
                return resolve(response);
            })
            .fail(error => {
                return reject(error);
            }); 
        });
    }
}