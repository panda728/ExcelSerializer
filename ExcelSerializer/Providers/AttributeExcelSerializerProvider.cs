﻿using System.Reflection;
using ExcelSerializer.Serializers;

namespace ExcelSerializer.Providers;

public sealed class AttributeExcelSerializerProvider : IExcelSerializerProvider
{
    public static IExcelSerializerProvider Instance { get; } = new AttributeExcelSerializerProvider();

    AttributeExcelSerializerProvider()
    {
    }

    public IExcelSerializer<T>? GetSerializer<T>() => Cache<T>.Serializer;

    static IExcelSerializer? CreateSerializer(Type type)
    {
        try
        {
            var attr = type.GetCustomAttribute<ExcelSerializerAttribute>();
            if (attr != null)
            {
                attr.Validate(type);
                return (IExcelSerializer?)Activator.CreateInstance(attr.Type);
            }

            return null;
        }
        catch (Exception ex)
        {
            return ErrorSerializer.Create(type, ex);
        }
    }

    private static class Cache<T>
    {
        public static readonly IExcelSerializer<T>? Serializer = (IExcelSerializer<T>?)CreateSerializer(typeof(T));
    }
}