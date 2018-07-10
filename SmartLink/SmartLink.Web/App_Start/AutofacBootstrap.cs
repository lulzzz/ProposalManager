// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using Autofac;
using Autofac.Core;
using AutoMapper;
using SmartLink.Service;
using SmartLink.Web.Mappings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SmartLink.Web
{
    public  class AutofacBootstrap
    {
        internal static void Init(ContainerBuilder builder)
        {
            builder.RegisterType<SourceService>().As<ISourceService>().InstancePerRequest();
            builder.RegisterType<DestinationService>().As<IDestinationService>().InstancePerRequest();
            builder.RegisterType<RecentFileService>().As<IRecentFileService>().InstancePerRequest();
            builder.RegisterType<SmartlinkDbContext>().AsSelf().InstancePerRequest();
            builder.RegisterType<ConfigService>().As<IConfigService>().SingleInstance();
            builder.RegisterType<AzureStorageService>().As<IAzureStorageService>().SingleInstance();
            builder.RegisterType<LogService>().As<ILogService>().SingleInstance();
            builder.RegisterType<MailService>().As<IMailService>().SingleInstance();
            builder.RegisterType<UserProfileService>().As<IUserProfileService>().InstancePerRequest();

            var mapperConfiguration = new MapperConfiguration(cfg =>
            {
                cfg.AddProfile(new MappingProfile());
                cfg.AddProfile(new ServiceMappingProfile());
                //This list is keep on going...

            });
            var mapper = mapperConfiguration.CreateMapper();
            builder.RegisterInstance(mapper).As<IMapper>().SingleInstance();
        }
    }
}