using System;

public class EmailTemplate
{
    public Guid EmailTemplateId { get; set; }
    public string Subject { get; set; }
    public string Body { get; set; }
    public EmailTemplate()
    {
        EmailTemplateId = Guid.Empty;
        Subject = string.Empty;
        Body = string.Empty;
    }
}