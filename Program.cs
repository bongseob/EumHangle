using System;
using System.IO;
using HwpReportGenerator.Core;
using HwpReportGenerator.Models;
using HwpReportGenerator.Helpers;

namespace HwpReportGenerator;

public class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("HWP 보고서 생성을 시작합니다.");

        // 1. 생성할 문서의 내용을 모델 객체로 정의
        var doc = new HwpDocument { Title = "월간 보고서" };

        var titleStyle = new HwpStyle { FontSize = 20, IsBold = true, Alignment = HwpAlignment.Center };
        doc.Add(new HwpParagraph("2025년 12월 월간 업무 보고서", titleStyle));
        
        doc.Add(new HwpParagraph("")); // 빈 줄

        var bodyStyle = new HwpStyle { FontSize = 12 };
        doc.Add(new HwpParagraph("1. 주요 성과", bodyStyle));
        
        // 표 생성
        var table = new HwpTable(3, 3);
        table.SetCell(0, 0, "구분");
        table.SetCell(0, 1, "목표");
        table.SetCell(0, 2, "실적");
        table.SetCell(1, 0, "A 프로젝트");
        table.SetCell(1, 1, "100건");
        table.SetCell(1, 2, "120건");
        table.SetCell(2, 0, "B 프로젝트");
        table.SetCell(2, 1, "50건");
        table.SetCell(2, 2, "55건");
        doc.Add(table);

        doc.Add(new HwpParagraph("2. 참고 자료", bodyStyle));
        
        // 이미지 파일 경로 (프로그램 실행 위치에 temp 폴더 및 이미지가 있다고 가정)
        string imagePath = Path.Combine(AppContext.BaseDirectory, "chart.png");
        if (File.Exists(imagePath))
        {
            doc.Add(new HwpImage(imagePath));
        }
        else
        {
            doc.Add(new HwpParagraph($