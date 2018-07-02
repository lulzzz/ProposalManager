// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using AutoMapper;
using SmartLink.Entity;
using SmartLink.Web.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SmartLink.Web.Mappings
{
    public class MappingProfile : Profile
    {
        public override string ProfileName
        {
            get
            {
                return "DomainViewModelMappings";
            }
        }

        public MappingProfile()
        {
            CreateMap<SourcePointForm, SourcePoint>();
            CreateMap<DestinationPointForm, DestinationPoint>()
                .ForMember(dest => dest.ReferencedSourcePoint, opt => opt.MapFrom(source => new SourcePoint() { Id = Guid.Parse(source.SourcePointId) }))
                .ForMember(dest => dest.CustomFormats, opt => opt.MapFrom(source => source.CustomFormatIds.Select(o => new CustomFormat() { Id = o })));
        }
    }
}