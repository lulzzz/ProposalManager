export class LocalizationService
{
    public getString(key: string)
    {
        return (window as any).stringResources[key];
    }
}